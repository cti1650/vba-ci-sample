Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-Path([string]$Path) {
  if (-not (Test-Path $Path)) { throw "Path not found: $Path" }
}

# Paths
$repoRoot = (Resolve-Path ".").Path
$vbaRoot  = Join-Path $repoRoot "vba"
$srcDir   = Join-Path $vbaRoot "src"
$testDir  = Join-Path $vbaRoot "test"
$runnerDir= Join-Path $vbaRoot "runner"

Assert-Path $srcDir
Assert-Path $testDir
Assert-Path $runnerDir

# Clean old markers
$success = "C:\temp\success.txt"
$error   = "C:\temp\error.txt"
$outXlsm = "C:\temp\vba_test.xlsm"

Remove-Item -Force -ErrorAction SilentlyContinue $success, $error, $outXlsm | Out-Null

# NOTE:
# Importing modules requires "Trust access to the VBA project object model" enabled.
# On GitHub Actions hosted runners this is typically enabled / permitted.
# If this script fails at VBProject access, see README for workaround.

$excel = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $wb = $excel.Workbooks.Add()
  $vbproj = $wb.VBProject

  # Import all .bas and .cls from src/test/runner
  Get-ChildItem -Path $srcDir -File -Include *.bas, *.cls | ForEach-Object {
    Write-Host "Importing src: $($_.FullName)"
    [void]$vbproj.VBComponents.Import($_.FullName)
  }

  Get-ChildItem -Path $testDir -File -Include *.bas, *.cls | ForEach-Object {
    Write-Host "Importing test: $($_.FullName)"
    [void]$vbproj.VBComponents.Import($_.FullName)
  }

  Get-ChildItem -Path $runnerDir -File -Include *.bas, *.cls | ForEach-Object {
    Write-Host "Importing runner: $($_.FullName)"
    [void]$vbproj.VBComponents.Import($_.FullName)
  }

  # Save as xlsm (52)
  Write-Host "Saving workbook: $outXlsm"
  $wb.SaveAs($outXlsm, 52)

  # Run test entrypoint
  Write-Host "Running VBA entrypoint: RunAll"
  $excel.Run("RunAll")

  # Close
  $wb.Close($false)
  $excel.Quit()

  # Decide result by marker files
  if (Test-Path $error) {
    Write-Host "=== VBA tests FAILED ==="
    Get-Content $error | Write-Host
    exit 1
  }

  if (-not (Test-Path $success)) {
    throw "No success marker found. Tests may not have executed. (Expected $success)"
  }

  Write-Host "=== VBA tests PASSED ==="
  Get-Content $success | Write-Host
  exit 0
}
catch {
  # Try to write error marker for artifact
  try {
    $msg = $_.Exception.Message
    $stack = $_.ScriptStackTrace
    "PSERROR: $msg`n$stack" | Out-File -FilePath $error -Encoding utf8
  } catch {}
  throw
}
finally {
  # Ensure Excel process ends
  if ($excel -ne $null) {
    try { $excel.Quit() | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
