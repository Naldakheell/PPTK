<#
.SYNOPSIS
  Automated PowerPoint test for PPTK Align Matrix layout (ModulePPTKSmartAlign: PPTK_AlignMatrixSpec,
  PPTK_MatrixBegin, PPTK_MatrixPreview, PPTK_MatrixRestore).

.DESCRIPTION
  Same rules as Test-Align.ps1: a private PowerPoint, a new unsaved presentation, the PPTM given by -Pptm opened
  without a window, macros called by qualified name through Application.Run (InvokeMember). The interactive grid
  (PPTKMatrixForm) is modal and is checked by hand; here the layout it asks for is applied directly as
  "rows|cols". Shapes are 40 x 30 pt, 20 pt apart, so the gap used is 20 pt. Exit code 0 when every case passes.
#>
param([string]$Pptm = '')

$ErrorActionPreference = 'Stop'
if (-not $Pptm) { $Pptm = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path 'PPTK-WINDOWS-FINAL.pptm' }
$Pptm = (Resolve-Path -LiteralPath $Pptm).Path
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Write-Host 'SKIPPED: close PowerPoint first'; exit 1 }
$results = New-Object System.Collections.ArrayList
function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) { [void]$results.Add($ok); Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail) }
function Near([double]$a, [double]$b) { [Math]::Abs($a - $b) -lt 0.05 }

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$addin = $null; $pres = $null
try {
    $addin = $app.Presentations.Open($Pptm, -1, 0, 0)
    $prefix = "'" + $addin.Name + "'!"
    function Run([string]$macro, [string]$arg = $null) {
        $a = if ($arg) { [object[]]@(($prefix + $macro), $arg) } else { [object[]]@(($prefix + $macro)) }
        return [System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, $a)
    }
    $pres = $app.Presentations.Add(-1)
    $s = $pres.Slides.Add(1, 12)
    $app.ActiveWindow.ViewType = 9
    $slideW = $pres.PageSetup.SlideWidth; $slideH = $pres.PageSetup.SlideHeight

    # n shapes in one row from (x, y), 40 x 30, 20 pt apart; names S1..Sn in reading order. All selected.
    function Row([int]$n, [double]$x, [double]$y, [int]$perRow = 0) {
        while ($s.Shapes.Count -gt 0) { $s.Shapes.Item(1).Delete() }
        if ($perRow -lt 1) { $perRow = $n }
        for ($i = 0; $i -lt $n; $i++) { $sh = $s.Shapes.AddShape(1, $x + ($i % $perRow) * 60, $y + [Math]::Floor($i / $perRow) * 50, 40, 30); $sh.Name = "S$($i + 1)" }
        $s.Shapes.Range([object[]]@(1..$n | ForEach-Object { "S$_" })).Select()
    }
    function Pos([int]$n) { @(1..$n | ForEach-Object { $x = $s.Shapes.Item("S$_"); "$([Math]::Round($x.Left,2)),$([Math]::Round($x.Top,2))" }) -join ';' }
    # Expected: cell (r, c) of a grid from (x, y) with 40 x 30 cells and the given gap.
    function Grid([int]$n, [int]$cols, [double]$x, [double]$y, [double]$gap = 20) { @(for ($i = 0; $i -lt $n; $i++) { "$([Math]::Round($x + ($i % $cols) * (40 + $gap),2)),$([Math]::Round($y + [Math]::Floor($i / $cols) * (30 + $gap),2))" }) -join ';' }
    function Case([string]$id, [string]$label, [int]$n, [string]$spec, [int]$cols, [int]$perRow = 0) {
        Row $n 100 100 $perRow
        $count = [int](Run 'ModulePPTKSmartAlign.PPTK_AlignMatrixSpec' $spec)
        $got = Pos $n; $want = Grid $n $cols 100 100
        Record $id $label ($count -eq $n -and $got -eq $want) ("spec $spec; got " + $got)
    }

    Case 'G2' '2 shapes, 1 x 2' 2 '1|2' 2
    Case 'G6' '6 shapes, 4 columns x 2 rows (2 empty cells), from the selection corner' 6 '2|4' 4
    Case 'G6a' '6 shapes, row 1 under the mouse: rows raised to 2' 6 '1|4' 4
    Case 'G6b' '6 shapes in 4 x 4: extra empty rows do not move anything' 6 '4|4' 4
    Case 'G7' '7 shapes, 3 x 3: last row has one shape on the left' 7 '3|3' 3
    Case 'G13' '13 shapes in one row (more than the old limit of 12)' 13 '1|13' 13 7
    Case 'G25' '25 shapes, 5 x 5' 25 '5|5' 5 5
    Case 'G25b' '25 shapes, 3 rows x 9 columns' 25 '3|9' 9 5

    # Reading order: shapes created out of order are placed by position, not by creation.
    Row 4 100 100
    $s.Shapes.Item('S1').Left = 400; $s.Shapes.Item('S1').Top = 300   # S1 is now last in reading order
    $s.Shapes.Range([object[]]@('S1', 'S2', 'S3', 'S4')).Select()
    Run 'ModulePPTKSmartAlign.PPTK_AlignMatrixSpec' '2|2' | Out-Null
    $ok = (Near $s.Shapes.Item('S2').Left 160) -and (Near $s.Shapes.Item('S2').Top 100) -and (Near $s.Shapes.Item('S1').Top 150) -and (Near $s.Shapes.Item('S1').Left 220)
    Record 'O1' 'reading order (top, then left), grid from the selection corner' $ok (Pos 4)

    # Near the right edge: the gap shrinks to 4, then the grid moves in to stay on the slide.
    Row 4 820 100
    Run 'ModulePPTKSmartAlign.PPTK_AlignMatrixSpec' '1|4' | Out-Null
    $last = $s.Shapes.Item('S4'); $gap = $s.Shapes.Item('S2').Left - ($s.Shapes.Item('S1').Left + 40)
    Record 'E1' 'grid stays on the slide (gap 4, moved in)' ((Near ($last.Left + $last.Width) ($slideW - 8)) -and (Near $gap 4)) ((Pos 4) + " gap $gap")

    # Sizes never change.
    Row 6 100 100; $s.Shapes.Item('S3').Width = 80; $s.Shapes.Item('S3').Height = 50
    $s.Shapes.Range([object[]]@(1..6 | ForEach-Object { "S$_" })).Select()
    Run 'ModulePPTKSmartAlign.PPTK_AlignMatrixSpec' '2|3' | Out-Null
    $sizes = @(1..6 | ForEach-Object { $x = $s.Shapes.Item("S$_"); "$($x.Width)x$($x.Height)" }) -join ' '
    $s3 = $s.Shapes.Item('S3'); $s1 = $s.Shapes.Item('S1')
    Record 'Z1' 'sizes kept; cells use the largest shape, smaller shapes centered' ($sizes -eq '40x30 40x30 80x50 40x30 40x30 40x30' -and (Near ($s1.Left + 20) 140) -and (Near $s3.Left 300)) ($sizes + ' | ' + (Pos 6))

    # Live preview, then Cancel: every shape back exactly.
    Row 6 100 100 3; $s.Shapes.Item('S5').Left = 333.3
    $s.Shapes.Range([object[]]@(1..6 | ForEach-Object { "S$_" })).Select()
    $before = Pos 6
    $n = [int](Run 'ModulePPTKSmartAlign.PPTK_MatrixBegin')
    Run 'ModulePPTKSmartAlign.PPTK_MatrixPreview' '1|6' | Out-Null; $p1 = Pos 6
    Run 'ModulePPTKSmartAlign.PPTK_MatrixPreview' '3|2' | Out-Null; $p2 = Pos 6
    Run 'ModulePPTKSmartAlign.PPTK_MatrixRestore' | Out-Null
    Record 'P1' 'preview moves the shapes; restore puts every one back' ($n -eq 6 -and $p1 -ne $before -and $p2 -ne $p1 -and (Pos 6) -eq $before) "before $before | restored $(Pos 6)"

    # One Undo after the final layout.
    Row 6 100 100 3; $before = Pos 6
    Run 'ModulePPTKSmartAlign.PPTK_AlignMatrixSpec' '1|6' | Out-Null; $after = Pos 6
    $app.CommandBars.ExecuteMso('Undo')
    Record 'U1' 'one Undo restores every shape' ($after -ne $before -and (Pos 6) -eq $before) "after $after | undone $(Pos 6)"

    # Fewer than two usable shapes: nothing happens.
    Row 1 100 100
    $count = [int](Run 'ModulePPTKSmartAlign.PPTK_AlignMatrixSpec' '1|1')
    Record 'N1' 'one shape: nothing to arrange' ($count -lt 2 -and (Pos 1) -eq '100,100') (Pos 1)
} catch {
    Record 'X' 'unexpected error' $false $_.Exception.Message
} finally {
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    try { if ($addin) { $addin.Close() } } catch { }
    try { $app.Visible = 0 } catch { }
    $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
$failed = @($results | Where-Object { -not $_ }).Count
Write-Host "TOTAL $($results.Count) tests, $failed failed"
exit $(if ($failed) { 1 } else { 0 })
