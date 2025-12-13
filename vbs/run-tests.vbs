' run-tests.vbs - テスト実行エントリーポイント
' VBAから変換されたVBSファイルを読み込み、Test_* 関数を自動検出して実行する
Option Explicit

Dim fso, scriptDir, genDir, compatPath
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
genDir = fso.BuildPath(scriptDir, "generated")
compatPath = fso.BuildPath(scriptDir, "vba-compat.vbs")

' VBA互換レイヤーを読み込み
If fso.FileExists(compatPath) Then
    ExecuteGlobal fso.OpenTextFile(compatPath).ReadAll
End If

' 生成されたVBSファイルを全て読み込み
Dim file, files, code, fileContent, enumsPath
code = ""

If Not fso.FolderExists(genDir) Then
    WScript.Echo "ERROR: generated folder not found: " & genDir
    WScript.Quit 1
End If

' _enums.vbs を最初に読み込む（Enum定数の定義）
enumsPath = fso.BuildPath(genDir, "_enums.vbs")
If fso.FileExists(enumsPath) Then
    ExecuteGlobal fso.OpenTextFile(enumsPath).ReadAll
End If

' 各ファイルを個別にExecuteGlobalで実行
Set files = fso.GetFolder(genDir).Files
For Each file In files
    If LCase(fso.GetExtensionName(file.Name)) = "vbs" Then
        ' _enums.vbs は既に読み込み済みなのでスキップ
        If LCase(file.Name) = "_enums.vbs" Then
            ' Skip
        Else
            fileContent = fso.OpenTextFile(file.Path).ReadAll
            On Error Resume Next
            ExecuteGlobal fileContent
            If Err.Number <> 0 Then
                WScript.Echo "ERROR loading " & file.Name & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
            ' テスト検出用にコードを蓄積
            code = code & vbCrLf & fileContent
        End If
    End If
Next

' Test_ で始まる関数を検出して実行
Dim passCount, failCount, testNames
passCount = 0
failCount = 0

' コードからTest_で始まるSub/Functionを抽出
Set testNames = CreateObject("Scripting.Dictionary")
Dim regex, matches, match
Set regex = New RegExp
regex.Global = True
regex.IgnoreCase = True
regex.Pattern = "(?:Sub|Function)\s+(Test_\w+)\s*\("

Set matches = regex.Execute(code)
For Each match In matches
    testNames(match.SubMatches(0)) = True
Next

WScript.Echo "========================================="
WScript.Echo "VBS Test Runner"
WScript.Echo "========================================="
WScript.Echo ""

' 各テストを実行
Dim testName
For Each testName In testNames.Keys
    On Error Resume Next
    Execute "Call " & testName & "()"
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

' 結果サマリー
WScript.Echo ""
WScript.Echo "========================================="
WScript.Echo "Results: " & passCount & " passed, " & failCount & " failed"
WScript.Echo "========================================="

' テストが1つもない場合はエラー
If passCount + failCount = 0 Then
    WScript.Echo "ERROR: No tests found!"
    WScript.Quit 1
End If

' 終了コード
If failCount > 0 Then
    WScript.Quit 1
Else
    WScript.Quit 0
End If
