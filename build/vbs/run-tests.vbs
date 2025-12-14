' run-tests.vbs - テスト実行エントリーポイント
' VBAから変換されたVBSファイルを読み込み、Test_* 関数を自動検出して実行する
'
' 各Test_*関数を個別のVBSファイルとして生成し、cscriptで実行する
' これにより、1つのテストが失敗しても他のテストは続行できる
Option Explicit

Dim fso, shell, scriptDir, genDir, compatPath, tempDir
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
genDir = fso.BuildPath(scriptDir, "generated")
compatPath = fso.BuildPath(scriptDir, "vba-compat.vbs")
tempDir = fso.BuildPath(scriptDir, "temp_tests")

If Not fso.FolderExists(genDir) Then
    WScript.Echo "ERROR: generated folder not found: " & genDir
    WScript.Quit 1
End If

' テンポラリディレクトリを作成
If fso.FolderExists(tempDir) Then
    fso.DeleteFolder tempDir, True
End If
fso.CreateFolder tempDir

' ============================================
' Step 1: 共通コードを収集
' ============================================
Dim baseCode, file, files, enumsPath

WScript.Echo "--- Collecting base code ---"

' GetScriptDir関数
baseCode = "' === GetScriptDir ===" & vbCrLf & _
          "Function GetScriptDir()" & vbCrLf & _
          "    GetScriptDir = """ & scriptDir & """" & vbCrLf & _
          "End Function" & vbCrLf & vbCrLf

' vba-compat.vbs
If fso.FileExists(compatPath) Then
    If fso.GetFile(compatPath).Size > 0 Then
        baseCode = baseCode & "' === vba-compat.vbs ===" & vbCrLf & _
                  fso.OpenTextFile(compatPath).ReadAll & vbCrLf & vbCrLf
        WScript.Echo "[OK] vba-compat.vbs"
    End If
End If

' _enums.vbs
enumsPath = fso.BuildPath(genDir, "_enums.vbs")
If fso.FileExists(enumsPath) Then
    If fso.GetFile(enumsPath).Size > 0 Then
        baseCode = baseCode & "' === _enums.vbs ===" & vbCrLf & _
                  fso.OpenTextFile(enumsPath).ReadAll & vbCrLf & vbCrLf
        WScript.Echo "[OK] _enums.vbs"
    End If
End If

' 生成されたVBSファイル（クラス・モジュール）
Dim classCode
classCode = ""
Set files = fso.GetFolder(genDir).Files
For Each file In files
    If LCase(fso.GetExtensionName(file.Name)) = "vbs" Then
        If LCase(file.Name) <> "_enums.vbs" Then
            If file.Size > 0 Then
                classCode = classCode & "' === " & file.Name & " ===" & vbCrLf & _
                           fso.OpenTextFile(file.Path).ReadAll & vbCrLf & vbCrLf
                WScript.Echo "[OK] " & file.Name
            End If
        End If
    End If
Next

WScript.Echo "--- Base code collected ---"
WScript.Echo ""

' ============================================
' Step 2: Test_*関数を検出
' ============================================
Dim allCode, testFunctions, regex, matches, match
allCode = baseCode & classCode

Set testFunctions = CreateObject("Scripting.Dictionary")
Set regex = New RegExp
regex.Global = True
regex.IgnoreCase = True
regex.Pattern = "(?:Public\s+)?(?:Sub|Function)\s+(Test_\w+)\s*\("

Set matches = regex.Execute(allCode)
For Each match In matches
    testFunctions(match.SubMatches(0)) = True
Next

WScript.Echo "Found " & testFunctions.Count & " test function(s)"
WScript.Echo ""

If testFunctions.Count = 0 Then
    WScript.Echo "ERROR: No tests found!"
    ' テンポラリディレクトリを削除
    If fso.FolderExists(tempDir) Then
        fso.DeleteFolder tempDir, True
    End If
    WScript.Quit 1
End If

' ============================================
' Step 3: 各テストを個別に実行
' ============================================
WScript.Echo "========================================="
WScript.Echo "VBS Test Runner (Isolated Mode)"
WScript.Echo "========================================="
WScript.Echo ""

Dim testName, testCode, testFile, testPath, exitCode
Dim passCount, failCount, errorOutput
passCount = 0
failCount = 0

For Each testName In testFunctions.Keys
    ' テスト用VBSファイルを生成
    ' テストコード生成：WScript.Quitでテストを終了するため、On Error Resume Nextは使用しない
    testCode = baseCode & classCode & vbCrLf & _
               "' === Test Execution ===" & vbCrLf & _
               "Call " & testName & "()" & vbCrLf & _
               "WScript.Quit 0" & vbCrLf

    testPath = fso.BuildPath(tempDir, testName & ".vbs")
    Set testFile = fso.CreateTextFile(testPath, True)
    testFile.Write testCode
    testFile.Close

    ' デバッグ: テストファイルの内容を出力（Test_UtilsFailWorksのみ）
    If testName = "Test_UtilsFailWorks" Then
        WScript.Echo "=== DEBUG: Checking for UtilsFail in test file ==="
        Dim allContent
        allContent = fso.OpenTextFile(testPath).ReadAll
        If InStr(allContent, "Sub UtilsFail") > 0 Then
            WScript.Echo "[OK] UtilsFail is defined in test file"
        Else
            WScript.Echo "[ERROR] UtilsFail is NOT defined in test file!"
        End If
        If InStr(allContent, "WScript.Quit 1") > 0 Then
            WScript.Echo "[OK] WScript.Quit 1 is in test file"
        Else
            WScript.Echo "[ERROR] WScript.Quit 1 is NOT in test file!"
        End If
        WScript.Echo "=== END DEBUG ==="
    End If

    ' テストを実行
    Dim exec, output
    Set exec = shell.Exec("cscript //nologo """ & testPath & """")

    ' 出力を収集
    output = ""
    Do While Not exec.StdOut.AtEndOfStream
        output = output & exec.StdOut.ReadLine() & vbCrLf
    Loop
    Do While Not exec.StdErr.AtEndOfStream
        output = output & exec.StdErr.ReadLine() & vbCrLf
    Loop

    ' 終了を待機
    Do While exec.Status = 0
        WScript.Sleep 50
    Loop

    exitCode = exec.ExitCode

    If exitCode = 0 Then
        WScript.Echo "[PASS] " & testName
        passCount = passCount + 1
    Else
        WScript.Echo "[FAIL] " & testName
        If Len(Trim(output)) > 0 Then
            WScript.Echo "       " & Replace(Trim(output), vbCrLf, vbCrLf & "       ")
        End If
        failCount = failCount + 1
    End If
Next

' ============================================
' Step 4: クリーンアップと結果サマリー
' ============================================
' テンポラリディレクトリを削除
If fso.FolderExists(tempDir) Then
    fso.DeleteFolder tempDir, True
End If

WScript.Echo ""
WScript.Echo "========================================="
WScript.Echo "Results: " & passCount & " passed, " & failCount & " failed"
WScript.Echo "========================================="

If failCount > 0 Then
    WScript.Quit 1
Else
    WScript.Quit 0
End If
