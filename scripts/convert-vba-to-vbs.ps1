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

.PARAMETER InputDir
    入力ディレクトリ (VBAファイルがあるディレクトリ)

.PARAMETER OutputDir
    出力ディレクトリ (VBSファイルを出力するディレクトリ)
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$InputDir,

    [Parameter(Mandatory=$true)]
    [string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Convert-VbaToVbs {
    param(
        [string]$Content,
        [string]$FileName,
        [bool]$IsClass
    )

    $lines = $Content -split "`r?`n"
    $result = @()
    $skipUntilEnd = $false
    $className = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

    $skipIfBlock = $false
    $skipEnumBlock = $false
    $skipTypeBlock = $false

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

        # New Collection, New Dictionary などの型付きNew → New のまま（VBSでも動く）
        # ただし、VBS では CreateObject を使う方が一般的

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

# .bas と .cls ファイルを処理
$files = Get-ChildItem -Path $InputDir -Include "*.bas", "*.cls" -Recurse

Write-Host "========================================="
Write-Host "VBA to VBS Converter"
Write-Host "========================================="
Write-Host ""

$convertedCount = 0

foreach ($file in $files) {
    $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8
    $isClass = $file.Extension -eq ".cls"
    $converted = Convert-VbaToVbs -Content $content -FileName $file.Name -IsClass $isClass

    # 出力ファイル名: .bas/.cls → .vbs
    $outputName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".vbs"
    $outputPath = Join-Path $OutputDir $outputName

    # BOMなしUTF-8で出力
    [System.IO.File]::WriteAllText($outputPath, $converted, [System.Text.UTF8Encoding]::new($false))

    Write-Host "[CONVERTED] $($file.Name) -> $outputName"
    $convertedCount++
}

Write-Host ""
Write-Host "========================================="
Write-Host "Converted $convertedCount file(s)"
Write-Host "========================================="
