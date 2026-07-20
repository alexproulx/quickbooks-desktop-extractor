<#
.SYNOPSIS
  Export the QuickBooks General Ledger one file per year, gently on QuickBooks.

.DESCRIPTION
  Runs QBExtract.py once per calendar year over [StartYear..EndYear]. Each run
  is its own QuickBooks session, extracts accounts + one year-granularity GL
  report (2 SDK calls), and writes a separate JSON file. A failed/hung year is
  isolated and reported so you can re-run just that year -- the other years are
  unaffected.

  QuickBooks must already be OPEN with the company file loaded (the reliable
  attached-session pattern). Do NOT pass --company-file.

.PARAMETER StartYear
  First fiscal year to export. Set this to the company's earliest year with data.

.PARAMETER EndYear
  Last year to export (default: current year).

.PARAMETER OutDir
  Folder for the per-year JSON files (default: .\exports).

.PARAMETER TimeoutSec
  Per-year hard timeout. If a year hangs past this, its process is killed and
  the year is marked FAILED so the loop continues (default: 1800 = 30 min).

.PARAMETER Granularity
  GL chunk size per year: year (default, 1 report/year, lightest on QB),
  quarter, or month (finer failure isolation, more SDK calls).

.EXAMPLE
  .\export_years.ps1 -StartYear 2015
  .\export_years.ps1 -StartYear 2018 -EndYear 2024 -Granularity quarter
#>
param(
    [int]$StartYear = 2015,
    [int]$EndYear   = (Get-Date).Year,
    [string]$OutDir = "exports",
    [int]$TimeoutSec = 1800,
    [ValidateSet("year", "quarter", "month")]
    [string]$Granularity = "year"
)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

Write-Host "Exporting GL per year: $StartYear..$EndYear  (granularity=$Granularity, timeout=${TimeoutSec}s)"
Write-Host "Make sure QuickBooks is OPEN with the company file loaded.`n"

$results = @()
for ($y = $StartYear; $y -le $EndYear; $y++) {
    $out = Join-Path $OutDir "qbgl_$y.json"
    Write-Host "=== $y  ->  $out ==="

    $args = @(
        "QBExtract.py",
        "--year", $y,
        "--gl-granularity", $Granularity,
        "--no-pause",
        "--output", $out
    )
    # Run in-process-group so we can kill a hung run. Python's own per-request
    # timeout (600s) usually trips first; this is a hard backstop.
    $proc = Start-Process -FilePath "python" -ArgumentList $args -NoNewWindow -PassThru
    if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
        Write-Warning "  $y exceeded ${TimeoutSec}s -- killing and moving on."
        try { $proc.Kill() } catch {}
        $results += [pscustomobject]@{ Year = $y; Status = "TIMEOUT"; File = $out }
        continue
    }

    if ($proc.ExitCode -eq 0 -and (Test-Path $out)) {
        $sizeMB = [math]::Round((Get-Item $out).Length / 1MB, 1)
        $results += [pscustomobject]@{ Year = $y; Status = "OK"; File = "$out (${sizeMB} MB)" }
    } else {
        $results += [pscustomobject]@{ Year = $y; Status = "FAILED (exit $($proc.ExitCode))"; File = $out }
    }
    Write-Host ""
}

Write-Host "`n===== SUMMARY ====="
$results | Format-Table -AutoSize

$failed = $results | Where-Object { $_.Status -ne "OK" }
if ($failed) {
    Write-Host "`nRe-run failed years individually, e.g.:"
    foreach ($f in $failed) {
        Write-Host "  python QBExtract.py --year $($f.Year) --gl-granularity $Granularity --no-pause --output $OutDir\qbgl_$($f.Year).json"
    }
    exit 1
} else {
    Write-Host "`nAll years exported to $OutDir\"
}
