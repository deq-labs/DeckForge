<#
.SYNOPSIS
  Builds the PowerPoint-authored fixture templates (fixtures/templates/ppt-*.pptx) by driving
  real PowerPoint through COM. Windows + PowerPoint (with Excel, for chart data) required.

.DESCRIPTION
  These files carry PowerPoint's real-world quirks (placeholders, theme inheritance, run splits,
  lang attributes, embedded chart workbooks with tables), unlike the pptxgenjs stand-ins.
  Three of them double as the Lite template packs (see template-packs/).

  Run:  powershell -ExecutionPolicy Bypass -File tools/powerpoint/build-templates.ps1
#>
param(
  [string]$OutDir = (Join-Path $PSScriptRoot '..\..\fixtures\templates'),
  # Build only these templates, e.g. -Only KpiScorecard,EdgeCases
  [string[]]$Only = @('OpsReview', 'KpiScorecard', 'ClientMonthly', 'PerSite', 'EdgeCases')
)
$ErrorActionPreference = 'Stop'
$OutDir = (Resolve-Path $OutDir).Path

# ---------------------------------------------------------------- constants
function Rgb([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }  # Office colors are BGR
$NAVY = Rgb 31 56 100
$BLUE = Rgb 46 117 182
$LIGHT = Rgb 242 244 247
$DARK = Rgb 38 38 38
$MID = Rgb 118 118 118
$WHITE = Rgb 255 255 255
$EN = [string][char]0x2013   # en dash
$DOT = [string][char]0x00B7  # middle dot
$TRUE_ = -1; $FALSE_ = 0     # msoTriState

$xlColumnClustered = 51; $xlColumnStacked = 52; $xlBarClustered = 57
$xlLineMarkers = 65; $xlPie = 5; $xlDoughnut = -4120; $xlColumns = 2

# ---------------------------------------------------------------- helpers
function Get-Layout($pres, [string]$name) {
  foreach ($l in $pres.SlideMaster.CustomLayouts) { if ($l.Name -eq $name) { return $l } }
  throw "layout '$name' not found"
}

function Add-Slide($pres, [string]$layout, [string]$title = $null) {
  $s = $pres.Slides.AddSlide($pres.Slides.Count + 1, (Get-Layout $pres $layout))
  if ($title -ne $null -and $s.Shapes.HasTitle) {
    $tr = $s.Shapes.Title.TextFrame.TextRange
    $tr.Text = $title
    $tr.Font.Size = 30; $tr.Font.Bold = $TRUE_; $tr.Font.Color.RGB = $NAVY
  }
  return $s
}

function Add-Text($slide, [string]$text, $l, $t, $w, $h, [int]$size = 16, [bool]$bold = $false, $color = $DARK, [int]$align = 1) {
  $s = $slide.Shapes.AddTextbox(1, $l, $t, $w, $h)
  $s.TextFrame.WordWrap = $TRUE_
  $tr = $s.TextFrame.TextRange
  $tr.Text = $text
  $tr.Font.Size = $size
  $tr.Font.Bold = $(if ($bold) { $TRUE_ } else { $FALSE_ })
  $tr.Font.Color.RGB = $color
  $tr.ParagraphFormat.Alignment = $align
  return $s
}

function Add-Footer($slide, [string]$text) {
  $null = Add-Text $slide $text 36 500 600 24 10 $false $MID
}

function Add-Notes($slide, [string]$text) {
  $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $text
}

# A KPI tile: label, big value, small detail line.
function Add-Tile($slide, $l, $t, $w, [string]$label, [string]$value, [string]$detail) {
  $bg = $slide.Shapes.AddShape(5, $l, $t, $w, 150)  # rounded rectangle
  $bg.Fill.ForeColor.RGB = $LIGHT; $bg.Line.Visible = $FALSE_
  $bg.Adjustments.Item(1) = 0.08
  $null = Add-Text $slide $label ($l + 12) ($t + 12) ($w - 24) 24 13 $true $MID
  $null = Add-Text $slide $value ($l + 12) ($t + 42) ($w - 24) 50 32 $true $NAVY
  if ($detail) { $null = Add-Text $slide $detail ($l + 12) ($t + 104) ($w - 24) 36 11 $false $MID }
}

function Add-Table($slide, [string]$name, [string[]]$headers, [object[]]$rows, $l, $t, $w, [double[]]$widths) {
  $shape = $slide.Shapes.AddTable($rows.Count + 1, $headers.Count, $l, $t, $w, 30 * ($rows.Count + 1))
  $shape.Name = $name
  $tbl = $shape.Table
  for ($c = 0; $c -lt $headers.Count; $c++) {
    $cell = $tbl.Cell(1, $c + 1).Shape.TextFrame.TextRange
    $cell.Text = $headers[$c]; $cell.Font.Size = 14; $cell.Font.Bold = $TRUE_
    if ($widths) { $tbl.Columns.Item($c + 1).Width = $widths[$c] }
  }
  for ($r = 0; $r -lt $rows.Count; $r++) {
    $row = $rows[$r]
    for ($c = 0; $c -lt $headers.Count; $c++) {
      $cell = $tbl.Cell($r + 2, $c + 1).Shape.TextFrame.TextRange
      $cell.Text = [string]$row[$c]; $cell.Font.Size = 13
      if ($row[$c] -is [double] -or $row[$c] -is [int]) { $cell.ParagraphFormat.Alignment = 3 }
    }
  }
  return $shape
}

# COM late binding trips over mixed-type Value2 assignments; Formula takes a string and Excel parses it.
function Set-Cell($ws, [int]$r, [int]$c, $v) {
  $text = if ($v -is [double] -or $v -is [int]) { ([double]$v).ToString([cultureinfo]::InvariantCulture) } else { [string]$v }
  $ws.Cells.Item($r, $c).Formula = $text
}

function Add-Chart($slide, [int]$type, [string]$name, $l, $t, $w, $h, [string]$catHeader, [string[]]$series, [object[]]$rows) {
  $shape = $slide.Shapes.AddChart2(-1, $type, $l, $t, $w, $h, $true)
  $shape.Name = $name
  if ($env:DECKFORGE_TRACE) { Write-Host "  chart $name created" }
  $chart = $shape.Chart
  $chart.ChartData.Activate()
  # The embedded workbook opens asynchronously in Excel; wait for it.
  $wb = $null
  for ($i = 0; $i -lt 40 -and -not $wb; $i++) {
    try { $wb = $chart.ChartData.Workbook; $null = $wb.Worksheets.Item(1) } catch { $wb = $null }
    if (-not $wb) { Start-Sleep -Milliseconds 250 }
  }
  if (-not $wb) { throw "chart data workbook for $name did not open" }
  $ws = $wb.Worksheets.Item(1)
  $ws.Range('A1:Z200').ClearContents()
  Set-Cell $ws 1 1 $catHeader
  for ($c = 0; $c -lt $series.Count; $c++) { Set-Cell $ws 1 ($c + 2) $series[$c] }
  for ($r = 0; $r -lt $rows.Count; $r++) {
    $row = $rows[$r]
    for ($c = 0; $c -lt $row.Count; $c++) { Set-Cell $ws ($r + 2) ($c + 1) $row[$c] }
  }
  $lastCol = [char](65 + $series.Count)
  $lastRow = $rows.Count + 1
  if ($ws.ListObjects.Count -gt 0) { $ws.ListObjects.Item(1).Resize($ws.Range("A1:$lastCol$lastRow")) }
  $chart.SetSourceData(("='Sheet1'!`$A`$1:`${0}`${1}" -f $lastCol, $lastRow), $xlColumns)
  if ($env:DECKFORGE_TRACE) { Write-Host "  chart $name data set" }
  $wb.Close()
  if ($env:DECKFORGE_TRACE) { Write-Host "  chart $name workbook closed" }
  $chart.HasTitle = $false
  if ($series.Count -gt 1 -or $type -eq $xlPie -or $type -eq $xlDoughnut) { $chart.HasLegend = $true; $chart.Legend.Position = -4107 }  # bottom
  return $shape
}

function Save-Deck($pres, [string]$file) {
  $path = Join-Path $OutDir $file
  if (Test-Path $path) { Remove-Item $path }
  $pres.SaveAs($path, 24)  # ppSaveAsOpenXMLPresentation
  $pres.Close()
  Write-Host "wrote $path"
}

# ---------------------------------------------------------------- templates

function Build-OpsReview($app) {
  $p = $app.Presentations.Add($TRUE_)
  $s = Add-Slide $p 'Title Slide'
  $s.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = '{{title}}'
  $s.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "{{site}} $DOT {{period_start}} $EN {{period_end}}"

  $s = Add-Slide $p 'Title Only' 'Key metrics'
  Add-Tile $s 36 130 210 'OEE' '{{oee}}' "{{oee.trend}} {{oee.change}} vs {{oee.previous}}"
  Add-Tile $s 262 130 210 'Output (units)' '{{output_units}}' "{{output_units.trend}} {{output_units.change_pct}}"
  Add-Tile $s 488 130 210 'Scrap rate' '{{scrap_rate}}' "{{scrap_rate.trend}} vs {{scrap_rate.previous}}"
  Add-Tile $s 714 130 210 'Downtime (h)' '{{downtime_hours}}' "{{downtime_hours.trend}} {{downtime_hours.change}} h"
  Add-Footer $s "{{site}} $DOT {{period_end}}"

  $s = Add-Slide $p 'Title Only' 'Downtime by cause'
  $null = Add-Table $s 'downtime' @('Line', 'Cause', 'Minutes', 'Share') @(
    @('Line 1', 'Changeover', 120, '38%'),
    @('Line 2', 'Jam', 45, '14%'),
    @('Line 3', 'Material shortage', 150, '48%')
  ) 36 120 888 @(160, 448, 140, 140)
  $null = Add-Text $s 'Total downtime: {{downtime_hours}} h' 36 440 600 30 16 $true $NAVY
  Add-Footer $s "{{site}} $DOT {{period_end}}"

  $s = Add-Slide $p 'Title Only' 'Downtime by line'
  $null = Add-Chart $s $xlColumnClustered 'downtime_by_line' 36 110 888 380 'Line' @('Planned', 'Unplanned') @(
    @('Line 1', 60, 120), @('Line 2', 30, 45), @('Line 3', 45, 150), @('Line 4', 20, 35)
  )
  Add-Footer $s "{{site}} $DOT {{period_end}}"

  $s = Add-Slide $p 'Title Only' 'OEE trend'
  $c = Add-Chart $s $xlLineMarkers 'oee_trend' 36 110 888 380 'Month' @('OEE', 'Target') @(
    @('May', 0.71, 0.75), @('Jun', 0.74, 0.75), @('Jul', 0.69, 0.75), @('Aug', 0.76, 0.75), @('Sep', 0.78, 0.75), @('Oct', 0.8, 0.75)
  )
  $c.Chart.Axes(2).TickLabels.NumberFormat = '0%'
  Add-Footer $s "{{site}} $DOT {{period_end}}"

  $s = Add-Slide $p 'Title and Content' 'Highlights and actions'
  $s.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = '{{highlights}}'
  Add-Footer $s "{{site}} $DOT {{period_end}}"
  Save-Deck $p 'ppt-ops-review.pptx'
}

function Build-KpiScorecard($app) {
  $p = $app.Presentations.Add($TRUE_)
  $s = Add-Slide $p 'Title Slide'
  $s.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = '{{title}}'
  $s.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = '{{period}}'

  $s = Add-Slide $p 'Title Only' 'Scorecard'
  $null = Add-Table $s 'scorecard' @('Metric', 'Target', 'Actual', 'Status') @(
    @('On-time delivery', '95%', '97%', 'On track'),
    @('First pass yield', '98%', '96.5%', 'Watch'),
    @('Scrap rate', '2.0%', '1.8%', 'On track'),
    @('Customer complaints', '3', '5', 'Off track'),
    @('Safety incidents', '0', '0', 'On track')
  ) 36 120 888 @(368, 160, 160, 200)
  $null = Add-Text $s 'Overall: {{overall_status}}' 36 440 600 30 16 $true $NAVY

  $s = Add-Slide $p 'Title Only' 'Scrap by cause'
  $null = Add-Chart $s $xlPie 'scrap_by_cause' 180 100 600 400 'Cause' @('Scrap (kg)') @(
    @('Setup', 420), @('Material', 310), @('Operator', 150), @('Machine', 260), @('Other', 60)
  )

  $s = Add-Slide $p 'Title Only' 'Volume vs yield'
  $c = Add-Chart $s $xlColumnClustered 'volume_vs_yield' 36 110 888 380 'Month' @('Volume', 'Yield') @(
    @('Jul', 12000, 0.962), @('Aug', 13400, 0.958), @('Sep', 12800, 0.967), @('Oct', 14100, 0.971)
  )
  $ser = $c.Chart.SeriesCollection(2)
  $ser.ChartType = $xlLineMarkers
  $ser.AxisGroup = 2
  $c.Chart.Axes(2, 2).TickLabels.NumberFormat = '0%'

  $s = Add-Slide $p 'Title Only' 'Safety'
  $null = Add-Text $s 'Incidents this period: {{incidents}}' 36 140 888 40 24 $true $NAVY
  $null = Add-Text $s '{{incident_notes}}' 36 200 888 200 16
  Add-Notes $s "#show-if incidents`r(This slide is left out when incidents is empty or 0.)"
  Save-Deck $p 'ppt-kpi-scorecard.pptx'
}

function Build-ClientMonthly($app) {
  $p = $app.Presentations.Add($TRUE_)
  $s = Add-Slide $p 'Title Slide'
  $s.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = '{{client}}'
  $s.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Monthly performance report $DOT {{month}}`rPrepared by {{agency}}"

  $s = Add-Slide $p 'Title Only' 'At a glance'
  Add-Tile $s 36 130 210 'Sessions' '{{sessions}}' '{{sessions.trend}} {{sessions.change_pct}} vs last month'
  Add-Tile $s 262 130 210 'Leads' '{{leads}}' '{{leads.trend}} {{leads.change_pct}} vs last month'
  Add-Tile $s 488 130 210 'Conversion rate' '{{conversion_rate}}' '{{conversion_rate.trend}} vs {{conversion_rate.previous}}'
  Add-Tile $s 714 130 210 'Ad spend' '{{spend}}' '{{spend.trend}} {{spend.change}}'
  Add-Footer $s "{{client}} $DOT {{month}}"

  $s = Add-Slide $p 'Title Only' 'Leads by channel'
  $null = Add-Chart $s $xlBarClustered 'leads_by_channel' 36 110 888 380 'Channel' @('This month', 'Last month') @(
    @('Organic search', 120, 104), @('Paid search', 95, 88), @('Social', 40, 52), @('Email', 33, 30), @('Referral', 18, 15)
  )
  Add-Footer $s "{{client}} $DOT {{month}}"

  $s = Add-Slide $p 'Title Only' 'Campaigns'
  $null = Add-Table $s 'campaigns' @('Campaign', 'Spend', 'Leads', 'Cost per lead') @(
    @('Spring promo', '$4,200', 61, '$68.85'),
    @('Brand search', '$2,100', 54, '$38.89'),
    @('Retargeting', '$900', 17, '$52.94')
  ) 36 120 888 @(408, 160, 160, 160)
  Add-Footer $s "{{client}} $DOT {{month}}"

  $s = Add-Slide $p 'Title and Content' 'Notes and next steps'
  $s.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = '{{notes}}'
  Add-Footer $s "{{client}} $DOT {{month}}"
  Save-Deck $p 'ppt-client-monthly.pptx'
}

function Build-PerSite($app) {
  $p = $app.Presentations.Add($TRUE_)
  $s = Add-Slide $p 'Title Slide'
  $s.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = '{{title}}'
  $s.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = '{{period_end}}'

  $s = Add-Slide $p 'Title Only' '{{site}}'
  Add-Tile $s 36 130 280 'OEE' '{{oee}}' ''
  Add-Tile $s 340 130 280 'Output (units)' '{{output}}' ''
  Add-Tile $s 644 130 280 'Downtime (h)' '{{downtime_hours}}' ''
  $null = Add-Text $s 'Plant manager: {{manager}}' 36 320 888 30 16
  Add-Notes $s "#repeat sites`r(One copy of this slide per row of the 'sites' sheet.)"

  $s = Add-Slide $p 'Title and Content' 'Summary'
  $s.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = '{{summary}}'
  Save-Deck $p 'ppt-per-site.pptx'
}

function Build-EdgeCases($app) {
  $p = $app.Presentations.Add($TRUE_)

  # Slide 1: tags split by formatting and language changes inside the tag.
  $s = Add-Slide $p 'Title Slide'
  $s.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = '{{title}}'
  $sub = $s.Shapes.Placeholders.Item(2).TextFrame.TextRange
  $sub.Text = 'Revenue: {{revenue}} / Units: {{units}}'
  $sub.Characters(10, 5).Font.Color.RGB = $BLUE   # "{{rev" in blue, "enue}}" not
  $sub.Characters(32, 4).LanguageID = 2057          # part of "{{units}}" marked en-GB

  # Slide 2: a tag inside a group, and a tag inside an unnamed (non-slot) table.
  $s = Add-Slide $p 'Title Only' 'Groups and plain tables'
  $box = $s.Shapes.AddShape(1, 36, 120, 300, 100); $box.Name = 'Group Box'
  $txt = Add-Text $s 'Grouped: {{grouped_value}}' 46 150 280 40 18
  $txt.Name = 'Group Text'
  $null = $s.Shapes.Range([object[]]@('Group Box', 'Group Text')).Group()
  $t = $s.Shapes.AddTable(2, 2, 400, 120, 500, 80)
  $t.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = 'Owner'
  $t.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = 'Due'
  $t.Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = '{{owner}}'
  $t.Table.Cell(2, 2).Shape.TextFrame.TextRange.Text = '{{due_date}}'
  $null = Add-Text $s 'Not tags: {single} {{ }} {{bad tag}} {{x.unknown}}' 36 440 888 30 12 $false $MID

  # Slide 3: things DeckForge must refuse or ignore, with clear warnings.
  $s = Add-Slide $p 'Title Only' 'Unsupported shapes'
  $m = Add-Table $s 'merged' @('A', 'B', 'C') @(@('1', '2', '3'), @('4', '5', '6')) 36 120 420 @(140, 140, 140)
  $m.Table.Cell(2, 1).Merge($m.Table.Cell(2, 2))
  $null = Add-Chart $s $xlColumnClustered 'Chart with spaces' 480 110 440 180 'X' @('Y') @(@('a', 1), @('b', 2))
  $null = Add-Chart $s $xlColumnStacked 'stacked' 480 300 440 180 'X' @('Y', 'Z') @(@('a', 1, 2), @('b', 2, 1))
  Add-Notes $s '{{tag_in_notes}} is not scanned.'
  Save-Deck $p 'ppt-edge-cases.pptx'
}

# ---------------------------------------------------------------- main
$app = New-Object -ComObject PowerPoint.Application
try {
  $app.DisplayAlerts = 1  # ppAlertsNone: never block on a dialog
  foreach ($name in $Only) {
    # Office automation occasionally drops the connection (RPC_E_DISCONNECTED); one retry is enough in practice.
    for ($attempt = 1; $attempt -le 2; $attempt++) {
      try {
        Write-Host "building $name..."
        & "Build-$name" $app
        break
      } catch {
        if ($attempt -eq 2) { throw }
        Write-Host "  retrying $name after: $($_.Exception.Message)"
        foreach ($pres in @($app.Presentations)) { if ($pres.Path -eq '') { $pres.Saved = $TRUE_; $pres.Close() } }
      }
    }
  }
} finally {
  foreach ($pres in @($app.Presentations)) { $pres.Saved = $TRUE_; $pres.Close() }
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
