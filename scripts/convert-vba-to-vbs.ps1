<#
.SYNOPSIS
    VBAファイル (.bas, .cls) をVBScript (.vbs) に変換する

.DESCRIPTION
    以下の変換を行う:
    - VBA固有のヘッダー行を削除 (VERSION, BEGIN/END, Attribute)
    - 型宣言を削除 (As Long, As String, As Variant 等)
    - ByVal/ByRef の型宣言を削除
    - 関数戻り値の型宣言を削除
    - .cls ファイルは Class ... End Class で囲む
    - Enum/Type ブロックを削除
    - Debug.Print を WScript.Echo に変換
    - Static 変数宣言を Dim に変換
    - DefInt/DefLng 等を削除
    - On Error GoTo ラベル を On Error Resume Next に変換

.PARAMETER InputDirs
    入力ディレクトリ (VBAファイルがあるディレクトリ、複数指定可能)

.PARAMETER OutputDir
    出力ディレクトリ (VBSファイルを出力するディレクトリ)
#>

param(
    [Parameter(Mandatory=$true)]
    [string[]]$InputDirs,

    [Parameter(Mandatory=$true)]
    [string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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
        [hashtable]$AllEnums = @{}
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
            if ($line -match "^\s*(\w+)\s*=\s*(\d+)") {
                $enumDefinitions += "Const ${currentEnumName}_$($matches[1]) = $($matches[2])"
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

        # Optional引数のデフォルト値付き型宣言を削除: Optional ByRef key As String = "" → Optional ByRef key = ""
        $converted = $converted -replace "(\bOptional\s+(?:ByVal\s+|ByRef\s+)?\w+)\s+As\s+\w+(\s*=)", '$1$2'

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

        # New Collection → CreateCollection() (vba-compat.vbs のモック使用)
        $converted = $converted -replace "\bNew\s+Collection\b", "CreateCollection()"

        # ThisWorkbook.path → GetScriptDir() (vba-compat.vbs で提供)
        $converted = $converted -replace "\bThisWorkbook\.path\b", "GetScriptDir()"
        $converted = $converted -replace "\bThisWorkbook\.Path\b", "GetScriptDir()"

        # Left$, Mid$, Right$, Replace$, Trim$, LTrim$, RTrim$, UCase$, LCase$, Space$, String$ → $ なし版
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

        # EnumName.Member → EnumName_Member に変換
        # ファイル内のEnum定義を使って変換
        foreach ($enumDef in $enumDefinitions) {
            if ($enumDef -match "Const\s+(\w+)_(\w+)\s*=") {
                $enumName = $matches[1]
                $memberName = $matches[2]
                $converted = $converted -replace "\b${enumName}\.${memberName}\b", "${enumName}_${memberName}"
            }
        }

        # 全ファイルから収集したEnum定義を使って変換（他ファイルで定義されたEnum参照用）
        foreach ($enumName in $AllEnums.Keys) {
            foreach ($memberName in $AllEnums[$enumName].Keys) {
                $converted = $converted -replace "\b${enumName}\.${memberName}\b", "${enumName}_${memberName}"
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

    # Enum定数を先頭に追加
    $enumBlock = ""
    if ($enumDefinitions.Count -gt 0) {
        $enumBlock = ($enumDefinitions -join "`r`n") + "`r`n`r`n"
    }

    # .cls ファイルは Class で囲む
    if ($IsClass) {
        # クラスの場合、Enum定数はクラス外に出力
        return "${enumBlock}Class $className`r`n$body`r`nEnd Class"
    } else {
        return "${enumBlock}$body"
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

# Step 1: 全ディレクトリからEnum定義を収集
$allEnums = @{}
foreach ($dir in $InputDirs) {
    if (Test-Path $dir) {
        $dirEnums = Collect-EnumDefinitions -InputDir $dir
        foreach ($enumName in $dirEnums.Keys) {
            if (-not $allEnums.ContainsKey($enumName)) {
                $allEnums[$enumName] = @{}
            }
            foreach ($memberName in $dirEnums[$enumName].Keys) {
                $allEnums[$enumName][$memberName] = $dirEnums[$enumName][$memberName]
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
        $converted = Convert-VbaToVbs -Content $content -FileName $file.Name -IsClass $isClass -AllEnums $allEnums

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
