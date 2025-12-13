<#
.SYNOPSIS
    VBAファイル (.bas, .cls) をVBScript (.vbs) に変換する

.DESCRIPTION
    以下の変換を行う:
    - VBA固有のヘッダー行を削除 (VERSION, BEGIN/END, Attribute)
    - 型宣言を削除 (As Long, As String, As Variant 等)
    - ByVal/ByRef の型宣言を削除
    - 関数戻り値の型宣言を削除

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
        [string]$FileName
    )

    $lines = $Content -split "`r?`n"
    $result = @()
    $skipUntilEnd = $false

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

        # 型宣言を削除: As Long, As String, As Boolean, As Variant, As Integer, As Double, As Object, As Collection 等
        # パターン: As <型名> (行末または , または ) の前)
        $converted = $line -replace "\s+As\s+\w+(?=\s*[,\)\r\n]|$)", ""

        # 関数の戻り値型を削除: Function Foo() As Long → Function Foo()
        $converted = $converted -replace "\)\s+As\s+\w+\s*$", ")"

        # Dim x As Long → Dim x
        $converted = $converted -replace "(\bDim\s+\w+)\s+As\s+\w+", '$1'

        # Private/Public 変数宣言の型も削除
        $converted = $converted -replace "(\b(?:Private|Public)\s+\w+)\s+As\s+\w+", '$1'

        $result += $converted
    }

    # 先頭の空行を削除
    while ($result.Count -gt 0 -and $result[0] -match "^\s*$") {
        $result = $result[1..($result.Count - 1)]
    }

    return ($result -join "`r`n")
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
    $converted = Convert-VbaToVbs -Content $content -FileName $file.Name

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
