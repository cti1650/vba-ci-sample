<#
.SYNOPSIS
    VBAファイル (.bas, .cls) をVBScript (.vbs) に変換する

.DESCRIPTION
    VBA固有の構文をVBSで動作するように変換する。

    主な変換:
    - VBA固有のヘッダー行を削除 (VERSION, BEGIN/END, Attribute)
    - 型宣言を削除 (As Long, As String, As Variant 等)
    - ByVal/ByRef の型宣言を削除
    - 関数/Property戻り値の型宣言を削除
    - .cls ファイルは Class ... End Class で囲む
    - Enum/Type ブロックを削除し、グローバル変数として出力
    - Debug.Print を DebugPrint に変換
    - Static 変数宣言を Dim に変換
    - DefInt/DefLng 等を削除
    - On Error GoTo ラベル を On Error Resume Next に変換
    - Optional引数のデフォルト値を削除
    - With New ClassName を複数行に分割
    - ThisWorkbook.Path を GetScriptDir() に変換 (Class内はインライン展開)
    - Left$/Mid$/Right$ 等の型付き関数を通常版に変換
    - Chr$/ChrW$/Str$/Hex$/Oct$/Date$/Time$/Error$ なども通常版に変換
    - Class内ではEnum定数をリテラル値に置換
    - Utils.Method → UtilsMethod に変換
    - ParamArray を削除
    - Event/RaiseEvent を削除
    - Friend キーワードを削除
    - WithEvents を削除
    - Implements を削除
    - Declare文（API宣言）を削除
    - #If/#End If 条件コンパイルを削除
    - GoSub/Return をコメントアウト
    - Open/Close/Print#/Input#/Line Input# をコメントアウト
    - LSet/RSet に警告コメント追加
    - AddressOf をコメントアウト
    - Resume/Resume Next をコメントアウト

    VBS互換性の制約:
    - VBSのClassからはグローバル変数/関数にアクセスできない
    - VBSのカスタムClassはFor Eachに対応していない
    - VBSにはOptional引数のデフォルト値がない
    - VBSにはStatic変数がない
    - VBSにはラベルへのジャンプがない
    - VBSにはEnum/Type/Event/ParamArrayがない
    - VBSにはFriend/WithEvents/Implementsがない
    - VBSにはDeclare（API宣言）がない
    - VBSにはGoSub/Returnがない
    - VBSにはOpen/Close#文法のファイルI/Oがない（FSOを使う）
    - VBSにはLSet/RSetがない
    - VBSにはAddressOfがない

.PARAMETER InputDirs
    入力ディレクトリ (VBAファイルがあるディレクトリ、複数指定可能)

.PARAMETER OutputDir
    出力ディレクトリ (VBSファイルを出力するディレクトリ)

.PARAMETER UseMockCreateObject
    CreateObject呼び出しをCreateObjectMockに変換する（テスト用）
#>

param(
    [Parameter(Mandatory=$true)]
    [string[]]$InputDirs,

    [Parameter(Mandatory=$true)]
    [string]$OutputDir,

    [Parameter(Mandatory=$false)]
    [switch]$UseMockCreateObject = $false
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# WinAPI関数のうち、vba-compat.vbs でモック提供されているもの
$global:MockedApiFunctions = @(
    # kernel32.dll
    "GetTickCount", "GetTickCount64", "Sleep", "GetCurrentProcessId",
    "GetLastError", "SetLastError", "GetComputerNameA", "GetComputerNameW",
    "GetComputerName", "GetUserNameA", "GetUserNameW", "GetTempPathA", "GetTempPathW",
    "GetSystemDirectoryA", "GetSystemDirectoryW", "GetWindowsDirectoryA", "GetWindowsDirectoryW",
    "QueryPerformanceCounter", "QueryPerformanceFrequency", "CopyMemory", "RtlMoveMemory",
    # user32.dll
    "MessageBoxA", "MessageBoxW", "GetActiveWindow", "GetForegroundWindow",
    "FindWindowA", "FindWindowW", "GetWindowTextA", "GetWindowTextW",
    "SetWindowTextA", "SetWindowTextW", "GetCursorPos", "SetCursorPos",
    "GetAsyncKeyState", "GetKeyState", "SendMessageA", "SendMessageW",
    "PostMessageA", "PostMessageW",
    # shell32.dll
    "SHGetFolderPathA", "SHGetFolderPathW",
    # ole32.dll
    "CoCreateGuid",
    # advapi32.dll
    "RegOpenKeyExA", "RegQueryValueExA", "RegCloseKey"
)

# 全ファイルからDeclare文（API宣言）を収集する関数
function Collect-ApiDeclarations {
    param(
        [string]$InputDir
    )

    $apiDeclarations = @{}  # FunctionName -> @{ Lib = "kernel32", Alias = "...", OriginalLine = "..." }
    $files = Get-ChildItem -Path $InputDir -Include "*.bas", "*.cls" -Recurse

    foreach ($file in $files) {
        $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8
        $lines = $content -split "`r?`n"

        foreach ($line in $lines) {
            # Declare文をパース
            # 例: Private Declare Function GetTickCount Lib "kernel32" () As Long
            # 例: Private Declare PtrSafe Function Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
            if ($line -match "^\s*(Public\s+|Private\s+)?Declare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)\s+Lib\s+""([^""]+)""") {
                $funcName = $matches[4]
                $libName = $matches[5]
                $alias = ""

                # Alias があるかチェック
                if ($line -match "Alias\s+""([^""]+)""") {
                    $alias = $matches[1]
                }

                $apiDeclarations[$funcName] = @{
                    Lib = $libName
                    Alias = $alias
                    OriginalLine = $line.Trim()
                    File = $file.Name
                }
            }
        }
    }

    return $apiDeclarations
}

# 全ファイルからEnum定義を収集する関数
function Collect-EnumDefinitions {
    param(
        [string]$InputDir
    )

    $allEnums = @{}  # EnumName -> @{ MemberName -> Value }
    $files = Get-ChildItem -Path $InputDir -Include "*.bas", "*.cls" -Recurse

    foreach ($file in $files) {
        $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8
        $lines = $content -split "`r?`n"

        $currentEnumName = ""
        $inEnum = $false
        $autoValue = 0

        foreach ($line in $lines) {
            if ($line -match "^\s*(Public\s+|Private\s+)?Enum\s+(\w+)") {
                $inEnum = $true
                $currentEnumName = $matches[2]
                $autoValue = 0
                if (-not $allEnums.ContainsKey($currentEnumName)) {
                    $allEnums[$currentEnumName] = @{}
                }
                continue
            }
            if ($inEnum) {
                if ($line -match "^\s*End\s+Enum") {
                    $inEnum = $false
                    $currentEnumName = ""
                    continue
                }
                # Enumメンバーを解析: MemberName = Value
                if ($line -match "^\s*(\w+)\s*=\s*(-?\d+)") {
                    $memberName = $matches[1]
                    $value = [int]$matches[2]
                    $allEnums[$currentEnumName][$memberName] = $value
                    $autoValue = $value + 1
                }
                # 値なしの場合: 自動採番
                elseif ($line -match "^\s*(\w+)\s*$" -and $matches[1] -ne "") {
                    $memberName = $matches[1]
                    $allEnums[$currentEnumName][$memberName] = $autoValue
                    $autoValue++
                }
            }
        }
    }

    return $allEnums
}

function Convert-VbaToVbs {
    param(
        [string]$Content,
        [string]$FileName,
        [bool]$IsClass,
        [hashtable]$AllEnums = @{},
        [hashtable]$AllApiDeclarations = @{},
        [bool]$UseMockCreateObject = $false
    )

    $lines = $Content -split "`r?`n"
    $result = @()
    $skipUntilEnd = $false
    $className = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

    $skipIfBlock = $false
    $skipEnumBlock = $false
    $skipTypeBlock = $false

    # Enum定義を収集して定数に変換
    $enumDefinitions = @()
    $currentEnumName = ""
    $inEnum = $false

    foreach ($line in $lines) {
        if ($line -match "^\s*(Public\s+|Private\s+)?Enum\s+(\w+)") {
            $inEnum = $true
            $currentEnumName = $matches[2]
            continue
        }
        if ($inEnum) {
            if ($line -match "^\s*End\s+Enum") {
                $inEnum = $false
                $currentEnumName = ""
                continue
            }
            # Enumメンバーを解析: MemberName = Value または MemberName
            # ExecuteGlobalではConstが使えないため、Dim宣言+代入で定義
            if ($line -match "^\s*(\w+)\s*=\s*(\d+)") {
                $varName = "${currentEnumName}_$($matches[1])"
                $enumDefinitions += "Dim ${varName} : ${varName} = $($matches[2])"
            }
            elseif ($line -match "^\s*(\w+)\s*$") {
                # 値なしの場合はスキップ（自動採番は複雑なので）
            }
        }
    }

    # 再度ループして変換
    $inEnum = $false
    foreach ($line in $lines) {
        # VERSION 1.0 CLASS ヘッダーブロックをスキップ
        if ($line -match "^VERSION\s+\d+\.\d+\s+CLASS") {
            $skipUntilEnd = $true
            continue
        }
        if ($skipUntilEnd) {
            if ($line -match "^END$") {
                $skipUntilEnd = $false
            }
            continue
        }

        # Attribute 行をスキップ
        if ($line -match "^\s*Attribute\s+VB_") {
            continue
        }

        # Option Explicit はスキップ（VBSでも使えるが、クラス内では不要）
        if ($line -match "^\s*Option\s+Explicit") {
            continue
        }

        # Option Compare, Option Base などもスキップ
        if ($line -match "^\s*Option\s+(Compare|Base|Private)") {
            continue
        }

        # #If VBA7 などの条件コンパイルブロックをスキップ
        if ($line -match "^\s*#If\s+") {
            $skipIfBlock = $true
            continue
        }
        if ($skipIfBlock) {
            if ($line -match "^\s*#End\s+If") {
                $skipIfBlock = $false
            }
            continue
        }

        # #Const もスキップ
        if ($line -match "^\s*#Const\s+") {
            continue
        }

        # Enum ブロックをスキップ (VBSにはEnumがない)
        if ($line -match "^\s*(Public\s+|Private\s+)?Enum\s+") {
            $skipEnumBlock = $true
            continue
        }
        if ($skipEnumBlock) {
            if ($line -match "^\s*End\s+Enum") {
                $skipEnumBlock = $false
            }
            continue
        }

        # Type ブロックをスキップ (VBSにはユーザー定義型がない)
        if ($line -match "^\s*(Public\s+|Private\s+)?Type\s+") {
            $skipTypeBlock = $true
            continue
        }
        if ($skipTypeBlock) {
            if ($line -match "^\s*End\s+Type") {
                $skipTypeBlock = $false
            }
            continue
        }

        # DefInt, DefLng, DefStr, DefBool, DefByte, DefCur, DefDate, DefDbl, DefSng, DefVar, DefObj をスキップ
        if ($line -match "^\s*Def(Int|Lng|Str|Bool|Byte|Cur|Date|Dbl|Sng|Var|Obj)\s+") {
            continue
        }

        # Implements をスキップ (VBSにはインターフェースがない)
        if ($line -match "^\s*Implements\s+") {
            continue
        }

        # Declare 文をスキップ (VBSにはAPI宣言がない)
        # 例: Private Declare Function GetTickCount Lib "kernel32" () As Long
        # 例: Private Declare PtrSafe Function ...
        if ($line -match "^\s*(Public\s+|Private\s+)?Declare\s+(PtrSafe\s+)?(Function|Sub)\s+") {
            continue
        }

        # GoSub/Return をコメントアウト (VBSにはGoSubがない)
        if ($line -match "^\s*GoSub\s+") {
            $line = "' [VBS UNSUPPORTED] " + $line
        }
        if ($line -match "^\s*Return\s*$") {
            $line = "' [VBS UNSUPPORTED] " + $line
        }

        # Line Input # をコメントアウト (VBSでは別の方法を使う)
        if ($line -match "^\s*Line\s+Input\s+#") {
            $line = "' [VBS UNSUPPORTED] " + $line
        }

        # Open ... For ... をコメントアウト (VBSではFSOを使う)
        if ($line -match "^\s*Open\s+.*\s+For\s+(Input|Output|Append|Binary|Random)\s+") {
            $line = "' [VBS UNSUPPORTED] " + $line
        }

        # Close # をコメントアウト
        if ($line -match "^\s*Close\s+#") {
            $line = "' [VBS UNSUPPORTED] " + $line
        }

        # Print #, Write #, Input # をコメントアウト
        if ($line -match "^\s*(Print|Write|Input)\s+#") {
            $line = "' [VBS UNSUPPORTED] " + $line
        }

        # Get #, Put # をコメントアウト (バイナリファイルI/O)
        if ($line -match "^\s*(Get|Put)\s+#") {
            $line = "' [VBS UNSUPPORTED] " + $line
        }

        # Seek # をコメントアウト
        if ($line -match "^\s*Seek\s+#") {
            $line = "' [VBS UNSUPPORTED] " + $line
        }

        # WithEvents をスキップ (VBSにはイベントがない)
        # 例: Private WithEvents obj As Object → Private obj
        $converted = $line -replace "\bWithEvents\s+", ""

        # Debug.Print → DebugPrint (vba-compat.vbs でモック提供)
        # VBSでは Print が予約語のため、関数呼び出しに変換
        $converted = $converted -replace "\bDebug\.Print\b", "DebugPrint"

        # モジュール名プレフィックスを変換 (VBSではモジュール名.関数名 で呼べない)
        # 例: Utils.Fail → UtilsFail, Utils.WriteTextFile → UtilsWriteTextFile
        $converted = $converted -replace "\bUtils\.", "Utils"

        # Static 変数 → Dim (VBSにはStaticがない)
        $converted = $converted -replace "^\s*Static\s+", "Dim "

        # On Error GoTo ラベル → On Error Resume Next (VBSではラベルジャンプ不可)
        # ただし On Error GoTo 0 はそのまま使える
        if ($converted -match "^\s*On\s+Error\s+GoTo\s+(?!0\s*$)") {
            $converted = $converted -replace "On\s+Error\s+GoTo\s+\w+", "On Error Resume Next"
        }

        # VBSではOptionalパラメータにデフォルト値を指定できない
        # 完全なパターン: Optional ByRef key As String = "" → key
        # 完全なパターン: Optional ByVal key As String = "" → ByVal key
        # 型宣言なし: Optional ByRef key = "" → key
        # 型宣言なし: Optional ByVal key = "" → ByVal key
        # ByRef/ByValなし: Optional key As String = "" → key
        # ByRef/ByValなし: Optional key = "" → key
        $converted = $converted -replace "\bOptional\s+ByRef\s+(\w+)\s+As\s+\w+\s*=\s*[^,\)]+", '$1'
        $converted = $converted -replace "\bOptional\s+ByVal\s+(\w+)\s+As\s+\w+\s*=\s*[^,\)]+", 'ByVal $1'
        $converted = $converted -replace "\bOptional\s+(\w+)\s+As\s+\w+\s*=\s*[^,\)]+", '$1'
        $converted = $converted -replace "\bOptional\s+ByRef\s+(\w+)\s*=\s*[^,\)]+", '$1'
        $converted = $converted -replace "\bOptional\s+ByVal\s+(\w+)\s*=\s*[^,\)]+", 'ByVal $1'
        $converted = $converted -replace "\bOptional\s+(\w+)\s*=\s*[^,\)]+", '$1'
        # デフォルト値なしのOptional: Optional ByRef key As String → key
        $converted = $converted -replace "\bOptional\s+ByRef\s+(\w+)\s+As\s+\w+", '$1'
        $converted = $converted -replace "\bOptional\s+ByVal\s+(\w+)\s+As\s+\w+", 'ByVal $1'
        $converted = $converted -replace "\bOptional\s+(\w+)\s+As\s+\w+", '$1'

        # 型宣言を削除: As Long, As String, As Boolean, As Variant, As Integer, As Double, As Object, As Collection 等
        $converted = $converted -replace "\s+As\s+\w+(?=\s*[,\)\r\n]|$)", ""

        # 関数の戻り値型を削除: Function Foo() As Long → Function Foo()
        $converted = $converted -replace "\)\s+As\s+\w+\s*$", ")"

        # Dim x As Long → Dim x
        $converted = $converted -replace "(\bDim\s+\w+)\s+As\s+\w+", '$1'

        # Private/Public 変数宣言の型も削除
        $converted = $converted -replace "(\b(?:Private|Public)\s+\w+)\s+As\s+\w+", '$1'

        # Const 宣言の型も削除: Const X As Long = 1 → Const X = 1
        $converted = $converted -replace "(\bConst\s+\w+)\s+As\s+\w+", '$1'

        # New Collection → VBSでの代替
        # Class内ではグローバル関数を呼べないため、直接Dictionaryを使う（Collectionモックは内部でDictionaryを使用）
        if ($IsClass) {
            # Class内ではDictionaryで代替（CollectionモックはFor Eachが使えないため）
            $converted = $converted -replace "\bNew\s+Collection\b", "CreateObject(""Scripting.Dictionary"")"
        } else {
            $converted = $converted -replace "\bNew\s+Collection\b", "CreateCollection()"
        }

        # With New ClassName → 複数行に分けて変換
        # VBSでは With New 構文がサポートされていない
        # また、: で複数ステートメントを繋ぐとWithで問題が起きる場合があるため、改行で分ける
        if ($converted -match "^(\s*)With\s+New\s+(\w+)") {
            $indent = $matches[1]
            $tempClassName = $matches[2]
            $converted = "${indent}Dim withTemp_$tempClassName`r`n${indent}Set withTemp_$tempClassName = New $tempClassName`r`n${indent}With withTemp_$tempClassName"
        }

        # ThisWorkbook.path → GetScriptDir() (vba-compat.vbs で提供)
        # ただし、Class内ではグローバル関数を呼べないため、インライン展開する
        if ($IsClass) {
            # Class内では直接FSO経由でパスを取得
            $converted = $converted -replace "\bThisWorkbook\.path\b", "CreateObject(""Scripting.FileSystemObject"").GetParentFolderName(WScript.ScriptFullName)"
            $converted = $converted -replace "\bThisWorkbook\.Path\b", "CreateObject(""Scripting.FileSystemObject"").GetParentFolderName(WScript.ScriptFullName)"
        } else {
            $converted = $converted -replace "\bThisWorkbook\.path\b", "GetScriptDir()"
            $converted = $converted -replace "\bThisWorkbook\.Path\b", "GetScriptDir()"
        }

        # Left$, Mid$, Right$, Replace$, Trim$, LTrim$, RTrim$, UCase$, LCase$, Space$, String$ → $ なし版
        # また、Chr$, ChrW$, ChrB$, Str$, Format$, Hex$, Oct$ も対応
        $converted = $converted -replace "\bLeft\$\(", "Left("
        $converted = $converted -replace "\bMid\$\(", "Mid("
        $converted = $converted -replace "\bRight\$\(", "Right("
        $converted = $converted -replace "\bReplace\$\(", "Replace("
        $converted = $converted -replace "\bTrim\$\(", "Trim("
        $converted = $converted -replace "\bLTrim\$\(", "LTrim("
        $converted = $converted -replace "\bRTrim\$\(", "RTrim("
        $converted = $converted -replace "\bUCase\$\(", "UCase("
        $converted = $converted -replace "\bLCase\$\(", "LCase("
        $converted = $converted -replace "\bSpace\$\(", "Space("
        $converted = $converted -replace "\bString\$\(", "String("
        $converted = $converted -replace "\bChr\$\(", "Chr("
        $converted = $converted -replace "\bChrW\$\(", "ChrW("
        $converted = $converted -replace "\bChrB\$\(", "Chr("
        $converted = $converted -replace "\bStr\$\(", "CStr("
        $converted = $converted -replace "\bFormat\$\(", "FormatNumber("
        $converted = $converted -replace "\bHex\$\(", "Hex("
        $converted = $converted -replace "\bOct\$\(", "Oct("
        $converted = $converted -replace "\bDate\$", "CStr(Date)"
        $converted = $converted -replace "\bTime\$", "CStr(Time)"
        $converted = $converted -replace "\bError\$\(", "Error("

        # LSet/RSet → 代替処理（VBSにはLSet/RSetがない）
        # LSet str = value → str = Left(value & Space(Len(str)), Len(str))
        # 完全な変換は複雑なのでコメントで警告
        if ($converted -match "\bLSet\s+") {
            $converted = "' [VBS WARNING: LSet needs manual conversion] " + $converted
        }
        if ($converted -match "\bRSet\s+") {
            $converted = "' [VBS WARNING: RSet needs manual conversion] " + $converted
        }

        # AddressOf をコメントアウト (VBSにはAddressOfがない)
        if ($converted -match "\bAddressOf\s+") {
            $converted = "' [VBS UNSUPPORTED: AddressOf] " + $converted
        }

        # Resume/Resume Next/Resume ラベル の処理
        # Resume Next は On Error Resume Next と組み合わせて使われるが、単独では問題
        if ($converted -match "^\s*Resume\s+Next\s*$") {
            # On Error Resume Next の後のResume Nextは通常不要
            $converted = "' [VBS: Resume Next removed] " + $converted
        }
        if ($converted -match "^\s*Resume\s+\w+\s*$" -and $converted -notmatch "^\s*Resume\s+Next") {
            $converted = "' [VBS UNSUPPORTED: Resume label] " + $converted
        }
        if ($converted -match "^\s*Resume\s*$") {
            $converted = "' [VBS UNSUPPORTED: Resume] " + $converted
        }

        # DoEvents → DoEvents() に変換（VBSでは引数なしでも括弧が必要な場合がある）
        # vba-compat.vbs でモック提供済み

        # Erase 配列 → 配列を空にリセット（VBSでも使える）
        # 変換不要

        # ParamArray を削除（VBSではParamArrayがない）
        # ParamArray args() As Variant → args
        $converted = $converted -replace "\bParamArray\s+(\w+)\(\)\s+As\s+\w+", '$1'
        $converted = $converted -replace "\bParamArray\s+(\w+)\(\)", '$1'

        # Event 宣言を削除（VBSにはEventがない）
        if ($converted -match "^\s*(Public\s+|Private\s+)?Event\s+") {
            continue
        }

        # RaiseEvent を削除（VBSにはRaiseEventがない）
        if ($converted -match "^\s*RaiseEvent\s+") {
            continue
        }

        # Friend を削除（VBSにはFriendがない）
        $converted = $converted -replace "\bFriend\s+(Sub|Function|Property)\b", '$1'

        # Property Get/Let/Set の型宣言を削除
        $converted = $converted -replace "(Property\s+(?:Get|Let|Set)\s+\w+\([^)]*\))\s+As\s+\w+", '$1'

        # EnumName.Member → EnumName_Member に変換
        # ファイル内のEnum定義を使って変換
        foreach ($enumDef in $enumDefinitions) {
            # "Dim EnumName_Member : EnumName_Member = Value" からEnum名とメンバー名を抽出
            if ($enumDef -match "Dim\s+(\w+)_(\w+)\s*:") {
                $enumName = $matches[1]
                $memberName = $matches[2]
                $converted = $converted -replace "\b${enumName}\.${memberName}\b", "${enumName}_${memberName}"
            }
        }

        # 全ファイルから収集したEnum定義を使って変換（他ファイルで定義されたEnum参照用）
        foreach ($enumName in $AllEnums.Keys) {
            foreach ($memberName in $AllEnums[$enumName].Keys) {
                # EnumName.Member → EnumName_Member
                $converted = $converted -replace "\b${enumName}\.${memberName}\b", "${enumName}_${memberName}"
            }
        }

        # Enumメンバー名だけの参照を変換（VBAでは同じEnum内ならEnum名を省略可能）
        # 例: Me.SetType = dictionary → Me.SetType = GlobDataType_dictionary
        # 代入の右辺 (= Member) パターンを検出
        foreach ($enumName in $AllEnums.Keys) {
            foreach ($memberName in $AllEnums[$enumName].Keys) {
                # = MemberName (代入の右辺でメンバー名のみ)
                $converted = $converted -replace "=\s*\b${memberName}\b(?!\s*[.(])", "= ${enumName}_${memberName}"
            }
        }

        # WinAPI関数呼び出しの処理
        # モック提供されていないAPI関数の呼び出しには警告コメントを追加
        foreach ($apiName in $AllApiDeclarations.Keys) {
            if ($converted -match "\b${apiName}\s*\(") {
                # モック提供されているかチェック
                if ($global:MockedApiFunctions -contains $apiName) {
                    # モック提供されている場合は何もしない（vba-compat.vbsで対応）
                } else {
                    # モック未提供のAPI呼び出しには警告を追加
                    $apiInfo = $AllApiDeclarations[$apiName]
                    if ($converted -notmatch "' \[VBS API MOCK\]") {
                        $converted = "' [VBS API MOCK: $apiName from $($apiInfo.Lib) - returns default value] " + $converted
                    }
                }
            }
        }

        # CreateObject をモック版に変換（オプション）
        if ($UseMockCreateObject) {
            # CreateObject("ProgID") → CreateObjectMock("ProgID")
            # ただし、既にCreateObjectMockの場合は変換しない
            # また、Scripting.FileSystemObject と Scripting.Dictionary は実際のオブジェクトを使う
            if ($converted -match 'CreateObject\s*\(\s*"([^"]+)"' -and $converted -notmatch 'CreateObjectMock') {
                $progId = $matches[1]
                # FSO と Dictionary は実際のオブジェクトを使う（基本機能として必要）
                if ($progId -notmatch "^Scripting\.(FileSystemObject|Dictionary)$") {
                    $converted = $converted -replace '\bCreateObject\s*\(', 'CreateObjectMock('
                }
            }
        }

        $result += $converted
    }

    # 先頭の空行を削除
    while ($result.Count -gt 0 -and $result[0] -match "^\s*$") {
        $result = $result[1..($result.Count - 1)]
    }

    # 末尾の空行を削除
    while ($result.Count -gt 0 -and $result[$result.Count - 1] -match "^\s*$") {
        $result = $result[0..($result.Count - 2)]
    }

    $body = $result -join "`r`n"

    # Enum定数は別ファイル(_enums.vbs)に出力
    # ただし、Class内ではグローバル変数にアクセスできないため、Class内ではリテラル値に置換する
    if ($IsClass) {
        foreach ($enumName in $AllEnums.Keys) {
            foreach ($memberName in $AllEnums[$enumName].Keys) {
                $value = $AllEnums[$enumName][$memberName]
                # EnumName_Member → リテラル値に置換
                $body = $body -replace "\b${enumName}_${memberName}\b", $value
            }
        }
    }

    # .cls ファイルは Class で囲む
    if ($IsClass) {
        return "Class $className`r`n$body`r`nEnd Class"
    } else {
        return $body
    }
}

# 出力ディレクトリを作成
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
}

Write-Host "========================================="
Write-Host "VBA to VBS Converter"
Write-Host "========================================="
Write-Host ""

# Step 1: 全ディレクトリからEnum定義とAPI宣言を収集
$allEnums = @{}
$allApiDeclarations = @{}

foreach ($dir in $InputDirs) {
    if (Test-Path $dir) {
        # Enum定義を収集
        $dirEnums = Collect-EnumDefinitions -InputDir $dir
        foreach ($enumName in $dirEnums.Keys) {
            if (-not $allEnums.ContainsKey($enumName)) {
                $allEnums[$enumName] = @{}
            }
            foreach ($memberName in $dirEnums[$enumName].Keys) {
                $allEnums[$enumName][$memberName] = $dirEnums[$enumName][$memberName]
            }
        }

        # API宣言を収集
        $dirApis = Collect-ApiDeclarations -InputDir $dir
        foreach ($apiName in $dirApis.Keys) {
            if (-not $allApiDeclarations.ContainsKey($apiName)) {
                $allApiDeclarations[$apiName] = $dirApis[$apiName]
            }
        }
    }
}

if ($allEnums.Count -gt 0) {
    Write-Host "[INFO] Collected Enums:"
    foreach ($enumName in $allEnums.Keys) {
        $members = ($allEnums[$enumName].Keys -join ", ")
        Write-Host "  - ${enumName}: $members"
    }
    Write-Host ""

    # Enum定数を _enums.vbs ファイルに出力（アンダースコアで始まるので最初に読み込まれる）
    $enumLines = @("' Auto-generated Enum constants")
    foreach ($enumName in $allEnums.Keys) {
        foreach ($memberName in $allEnums[$enumName].Keys) {
            $value = $allEnums[$enumName][$memberName]
            $enumLines += "${enumName}_${memberName} = ${value}"
        }
    }
    $enumContent = $enumLines -join "`r`n"
    $enumPath = Join-Path $OutputDir "_enums.vbs"
    [System.IO.File]::WriteAllText($enumPath, $enumContent, [System.Text.UTF8Encoding]::new($false))
    Write-Host "[GENERATED] _enums.vbs"
}

# API宣言情報を表示
if ($allApiDeclarations.Count -gt 0) {
    Write-Host "[INFO] Collected API Declarations:"
    $mockedCount = 0
    $unmockedCount = 0
    foreach ($apiName in $allApiDeclarations.Keys) {
        $apiInfo = $allApiDeclarations[$apiName]
        if ($global:MockedApiFunctions -contains $apiName) {
            Write-Host "  - ${apiName} (${($apiInfo.Lib)}) [MOCKED]"
            $mockedCount++
        } else {
            Write-Host "  - ${apiName} (${($apiInfo.Lib)}) [NOT MOCKED - will use default return]"
            $unmockedCount++
        }
    }
    Write-Host ""
    Write-Host "[INFO] API Summary: $mockedCount mocked, $unmockedCount not mocked"
    Write-Host ""
}

if ($UseMockCreateObject) {
    Write-Host "[INFO] CreateObject calls will be converted to CreateObjectMock"
    Write-Host ""
}

# Step 2: 全ディレクトリのファイルを変換
$convertedCount = 0

foreach ($dir in $InputDirs) {
    if (-not (Test-Path $dir)) {
        Write-Host "[WARN] Directory not found: $dir"
        continue
    }

    $files = Get-ChildItem -Path $dir -Include "*.bas", "*.cls" -Recurse

    foreach ($file in $files) {
        $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8
        $isClass = $file.Extension -eq ".cls"
        $converted = Convert-VbaToVbs -Content $content -FileName $file.Name -IsClass $isClass -AllEnums $allEnums -AllApiDeclarations $allApiDeclarations -UseMockCreateObject $UseMockCreateObject

        # 出力ファイル名: .bas/.cls → .vbs
        $outputName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".vbs"
        $outputPath = Join-Path $OutputDir $outputName

        # BOMなしUTF-8で出力
        [System.IO.File]::WriteAllText($outputPath, $converted, [System.Text.UTF8Encoding]::new($false))

        Write-Host "[CONVERTED] $($file.Name) -> $outputName"
        $convertedCount++
    }
}

Write-Host ""
Write-Host "========================================="
Write-Host "Converted $convertedCount file(s)"
Write-Host "========================================="
