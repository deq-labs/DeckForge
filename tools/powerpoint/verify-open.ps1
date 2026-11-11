<#
.SYNOPSIS
  Opens each .pptx in real PowerPoint (read-only, no window) and reports whether it opened
  cleanly. Optionally exports a PDF of each, which proves every slide renders.

.DESCRIPTION
  A file PowerPoint would "repair" either fails to open via automation or hangs on a prompt;
  each file runs in its own job with a timeout so one bad file can't block the run.
  Windows + PowerPoint required. Close other PowerPoint windows first: automation attaches to
  the running instance.

  Usage:
    powershell -ExecutionPolicy Bypass -File tools/powerpoint/verify-open.ps1 out/smoke/*.pptx [-Pdf]
#>
[CmdletBinding(PositionalBinding = $false)]
param(
  [Parameter(Mandatory = $true, Position = 0, ValueFromRemainingArguments = $true)][string[]]$Files,
  [switch]$Pdf,
  # Also export every slide as PNG into <deck>.slides/ (for eyeballing output).
  [switch]$Png,
  [int]$TimeoutSeconds = 90
)
$ErrorActionPreference = 'Stop'

$paths = foreach ($f in $Files) { Get-ChildItem -Path $f -File | ForEach-Object { $_.FullName } }
$failed = 0
foreach ($path in $paths) {
  $job = Start-Job -ArgumentList $path, $Pdf.IsPresent, $Png.IsPresent {
    param($path, $pdf, $png)
    $app = New-Object -ComObject PowerPoint.Application
    $app.DisplayAlerts = 1  # ppAlertsNone
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $pres = $app.Presentations.Open($path, -1, 0, 0)
    try {
      $count = $pres.Slides.Count
      $charts = 0
      foreach ($s in $pres.Slides) { foreach ($sh in $s.Shapes) { if ($sh.HasChart) { $charts++; $null = $sh.Chart.SeriesCollection().Count } } }
      if ($pdf) { $pres.SaveAs(($path -replace '\.pptx$', '.powerpoint.pdf'), 32) }  # ppSaveAsPDF
      if ($png) {
        $dir = $path -replace '\.pptx$', '.slides'
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $i = 1
        foreach ($s in $pres.Slides) { $s.Export((Join-Path $dir "slide$i.png"), 'PNG', 1280, 720); $i++ }
      }
      "$count slides, $charts charts"
    } finally {
      $pres.Close()
    }
  }
  if (Wait-Job $job -Timeout $TimeoutSeconds) {
    try {
      $info = Receive-Job $job -ErrorAction Stop
      Write-Host "ok   $path ($info)"
    } catch {
      $failed++
      Write-Host "FAIL $path : $($_.Exception.Message)"
    }
  } else {
    $failed++
    Write-Host "FAIL $path : no response after $TimeoutSeconds s (PowerPoint may be showing a repair prompt)"
    Stop-Job $job
  }
  Remove-Job $job -Force
}
if ($failed) { Write-Host "$failed file(s) did not open cleanly"; exit 1 }
Write-Host "all $($paths.Count) file(s) opened cleanly in PowerPoint"
