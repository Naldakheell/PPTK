<#
.SYNOPSIS
  Automated PowerPoint test for PPTK alignment (ModulePPTKAlign: the first selected shape is the reference)
  and for D H / D V (PPTK_DistributeLine: one row or column with even gaps; the ribbon buttons run the same code).

.DESCRIPTION
  Runs in a private PowerPoint instance (PowerPoint must be closed first), on a new unsaved presentation, with
  the PPTM given by -Pptm opened without a window. Macros are called with Application.Run and a qualified name,
  so the installed PPTK add-in (if it loads too) is never used. Shapes are selected in a known order with
  Shape.Select (checked on 2026-11-17 to give the same Selection.ShapeRange order as mouse clicks).
  The PPTK\Align\Reference setting is saved before and restored after. Nothing is saved to disk.
  Exit code 0 when every case passes.
#>
param([string]$Pptm = '')

$ErrorActionPreference = 'Stop'
if (-not $Pptm) { $Pptm = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path 'PPTK-WINDOWS-FINAL.pptm' }
$Pptm = (Resolve-Path -LiteralPath $Pptm).Path
# Never touch a PowerPoint the user is working in. An invisible instance with no presentations (an automation
# instance that has not exited yet) may be reused.
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    $idle = $false
    try { $running = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); $idle = ($running.Presentations.Count -eq 0 -and $running.Visible -eq 0) } catch { }
    if (-not $idle) { Write-Host 'SKIPPED: close PowerPoint first'; exit 1 }
}

$results = New-Object System.Collections.ArrayList
function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) { [void]$results.Add($ok); Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail) }
function Near([double]$a, [double]$b) { [Math]::Abs($a - $b) -lt 0.05 }

# Visible (rotated) bounding box, same math as ModuleObjects GetReal*.
function Box($s) {
    $r = (($s.Rotation % 360) + 360) % 360; $rad = $r * [Math]::PI / 180
    $cx = $s.Left + $s.Width / 2; $cy = $s.Top + $s.Height / 2
    if ($r -eq 0 -or $r -eq 180) { $w = $s.Width; $h = $s.Height }
    elseif ($r -eq 90 -or $r -eq 270) { $w = $s.Height; $h = $s.Width }
    else { $w = $s.Width * [Math]::Abs([Math]::Cos($rad)) + $s.Height * [Math]::Abs([Math]::Sin($rad)); $h = $s.Height * [Math]::Abs([Math]::Cos($rad)) + $s.Width * [Math]::Abs([Math]::Sin($rad)) }
    [pscustomobject]@{ L = $cx - $w / 2; T = $cy - $h / 2; R = $cx + $w / 2; B = $cy + $h / 2; C = $cx; M = $cy }
}

$regPath = 'HKCU:\Software\VB and VBA Program Settings\PPTK\Align'
$regBefore = (Get-ItemProperty -Path $regPath -Name Reference -ErrorAction SilentlyContinue).Reference

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$addin = $null; $pres = $null
try {
    $addin = $app.Presentations.Open($Pptm, -1, 0, 0)   # read-only, no window
    $prefix = "'" + $addin.Name + "'!"
    # Application.Run through InvokeMember: PowerShell cannot pass Run's ParamArray arguments directly.
    function Run([string]$macro, [string]$arg = $null) {
        $args2 = if ($null -ne $arg -and $arg -ne '') { [object[]]@(($prefix + $macro), $arg) } else { [object[]]@(($prefix + $macro)) }
        return [System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, $args2)
    }
    $pres = $app.Presentations.Add(-1)
    $slide = $pres.Slides.Add(1, 12)
    $app.ActiveWindow.ViewType = 9

    function NewShapes {
        while ($slide.Shapes.Count -gt 0) { $slide.Shapes.Item(1).Delete() }
        $a = $slide.Shapes.AddShape(1, 100, 100, 80, 60); $a.Name = 'A'
        $b = $slide.Shapes.AddShape(1, 300, 220, 120, 40); $b.Name = 'B'
        $c = $slide.Shapes.AddShape(1, 520, 50, 60, 100); $c.Name = 'C'
    }
    function SelectOrder([string[]]$names) { $slide.Shapes.Item($names[0]).Select(-1); for ($i = 1; $i -lt $names.Count; $i++) { $slide.Shapes.Item($names[$i]).Select(0) } }
    function Pos { @('A', 'B', 'C') | ForEach-Object { $s = $slide.Shapes.Item($_); "$_=$([Math]::Round($s.Left,2)),$([Math]::Round($s.Top,2))" } }

    # ---- 1. every edge, first selected = B (not leftmost/topmost), via the dispatcher token
    Run 'ModulePPTKAlign.PPTK_SetAlignReference' 'First'
    foreach ($edge in 'L', 'C', 'R', 'T', 'M', 'B') {
        NewShapes; SelectOrder @('B', 'C', 'A')
        $ref = Box ($slide.Shapes.Item('B')); $refLeft = $slide.Shapes.Item('B').Left; $refTop = $slide.Shapes.Item('B').Top
        Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' "A $edge"
        $okRef = (Near $slide.Shapes.Item('B').Left $refLeft) -and (Near $slide.Shapes.Item('B').Top $refTop)
        $okOthers = $true
        foreach ($n in 'A', 'C') { $bx = Box ($slide.Shapes.Item($n)); if (-not (Near $bx.$edge $ref.$edge)) { $okOthers = $false } }
        Record "E$edge" "A $edge aligns to the first selected shape, which stays" ($okRef -and $okOthers) ((Pos) -join ' ')
    }

    # ---- 2. rotated shapes: visible edges line up
    NewShapes; $slide.Shapes.Item('C').Rotation = 30; $slide.Shapes.Item('B').Rotation = 90
    SelectOrder @('B', 'A', 'C')
    $ref = Box ($slide.Shapes.Item('B'))
    Run 'ModulePPTKAlign.PPTK_AlignEdge' 'L'
    Record 'R1' 'rotated shapes: visible left edges match the rotated reference' ((Near (Box ($slide.Shapes.Item('A'))).L $ref.L) -and (Near (Box ($slide.Shapes.Item('C'))).L $ref.L)) ((Pos) -join ' ')
    foreach ($edge in 'T', 'B') {
        NewShapes; $slide.Shapes.Item('C').Rotation = 30; $slide.Shapes.Item('B').Rotation = 90
        SelectOrder @('B', 'A', 'C')
        $ref = Box ($slide.Shapes.Item('B'))
        Run 'ModulePPTKAlign.PPTK_AlignEdge' $edge
        Record "R$edge" "rotated shapes: visible $(if ($edge -eq 'T') { 'top' } else { 'bottom' }) edges match the first selected (rotated) shape" ((Near (Box ($slide.Shapes.Item('A'))).$edge $ref.$edge) -and (Near (Box ($slide.Shapes.Item('C'))).$edge $ref.$edge)) ((Pos) -join ' ')
    }

    # ---- 3. order matters: a different first shape gives a different result
    NewShapes; SelectOrder @('C', 'A', 'B')
    Run 'ModulePPTKAlign.PPTK_AlignEdge' 'T'
    Record 'O1' 'first selected C is the reference for A T' ((Near $slide.Shapes.Item('A').Top 50) -and (Near $slide.Shapes.Item('B').Top 50) -and (Near $slide.Shapes.Item('C').Top 50)) ((Pos) -join ' ')

    # ---- 4. inside a group: ChildShapeRange order
    NewShapes; $g = $slide.Shapes.Range([object[]]@('A', 'B', 'C')).Group(); $g.Name = 'G'
    $g.GroupItems.Item('C').Select(-1); $g.GroupItems.Item('A').Select(0)
    $cLeft = $g.GroupItems.Item('C').Left
    Run 'ModulePPTKAlign.PPTK_AlignEdge' 'L'
    Record 'G1' 'group children: aligned to the first selected child' ((Near $g.GroupItems.Item('A').Left $cLeft) -and (Near $g.GroupItems.Item('C').Left $cLeft) -and (Near $g.GroupItems.Item('B').Left 300)) ("A=$($g.GroupItems.Item('A').Left) C=$($g.GroupItems.Item('C').Left) B=$($g.GroupItems.Item('B').Left)")
    $g.Ungroup() | Out-Null

    # ---- 5. one shape: aligned to the slide (unchanged behavior)
    NewShapes; SelectOrder @('B')
    Run 'ModulePPTKAlign.PPTK_AlignEdge' 'L'
    Record 'S1' 'one shape aligns to the slide' (Near $slide.Shapes.Item('B').Left 0) ((Pos) -join ' ')

    # ---- 6. selection-bounds mode = PowerPoint Align (the previous behavior)
    Run 'ModulePPTKAlign.PPTK_SetAlignReference' 'Bounds'
    NewShapes; SelectOrder @('B', 'C', 'A')
    Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' 'A L'
    Record 'B1' 'bounds mode: all lefts at the leftmost shape (B moves)' ((Near $slide.Shapes.Item('A').Left 100) -and (Near $slide.Shapes.Item('B').Left 100) -and (Near $slide.Shapes.Item('C').Left 100)) ((Pos) -join ' ')
    $mode = Run 'ModulePPTKAlign.PPTK_AlignReference'
    Run 'ModulePPTKAlign.PPTK_SetAlignReference' 'First'
    $mode2 = Run 'ModulePPTKAlign.PPTK_AlignReference'
    Record 'B2' 'setting round trip Bounds -> First' ($mode -eq 'Bounds' -and $mode2 -eq 'First') "$mode -> $mode2"

    # ---- 7. one undo step restores every moved shape
    NewShapes; SelectOrder @('B', 'C', 'A'); $before = (Pos) -join ' '
    Run 'ModulePPTKAlign.PPTK_AlignEdge' 'R'
    $after = (Pos) -join ' '
    $app.CommandBars.ExecuteMso('Undo')
    Record 'U1' 'one Undo restores all shapes' ($after -ne $before -and ((Pos) -join ' ') -eq $before) "after: $after | undone: $((Pos) -join ' ')"

    # ---- 8. nothing selected / text selection: no error, nothing moves
    NewShapes; $app.ActiveWindow.Selection.Unselect(); $before = (Pos) -join ' '
    Run 'ModulePPTKAlign.PPTK_AlignEdge' 'L'
    Record 'N1' 'no selection: nothing happens' (((Pos) -join ' ') -eq $before) ''

    # ---- 9. D H / D V (PPTK_DistributeLine): one row or column, even gaps; ribbon and Smart Bar run the same code
    # Line check: every shape on the reference's middle (row) or center (column); gaps along the axis all equal
    # $gap (or all equal to each other when $gap is $null); the first shape along the axis starts at $start.
    function LineOk([string[]]$names, [string]$axis, [double]$cross, $start, $gap) {
        $boxes = @($names | ForEach-Object { Box ($slide.Shapes.Item($_)) })
        if ($axis -eq 'H') { $lo = 'L'; $hi = 'R'; $mid = 'M' } else { $lo = 'T'; $hi = 'B'; $mid = 'C' }
        $sorted = @($boxes | Sort-Object { $_.$lo })
        foreach ($b in $boxes) { if (-not (Near $b.$mid $cross)) { return $false } }
        if ($null -ne $start -and -not (Near $sorted[0].$lo $start)) { return $false }
        $first = $sorted[1].$lo - $sorted[0].$hi
        if ($null -ne $gap -and -not (Near $first $gap)) { return $false }
        if ($first -lt -0.05) { return $false }
        for ($i = 2; $i -lt $sorted.Count; $i++) { if (-not (Near ($sorted[$i].$lo - $sorted[$i - 1].$hi) $first)) { return $false } }
        return $true
    }
    function Stack([string]$axis) {
        while ($slide.Shapes.Count -gt 0) { $slide.Shapes.Item(1).Delete() }
        foreach ($k in 0..2) {
            if ($axis -eq 'H') { $x = $slide.Shapes.AddShape(1, 405, 220 + 12 * $k, 150, 100) } else { $x = $slide.Shapes.AddShape(1, 300 + 12 * $k, 200, 150, 100) }
            $x.Name = [string][char](65 + $k)
        }
    }
    Run 'ModulePPTKAlign.PPTK_SetAlignReference' 'First'

    # Scattered A B C, first selected A: the row is on A's middle, from the leftmost (A, 100) to the rightmost (C, 580).
    NewShapes; SelectOrder @('A', 'B', 'C'); $m = (Box ($slide.Shapes.Item('A'))).M
    Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' 'D H'
    Record 'DH1' 'D H: one row on the first selected, even gaps, outer edges kept' ((LineOk @('A', 'B', 'C') 'H' $m 100 $null) -and (Near (Box ($slide.Shapes.Item('C'))).R 580)) ((Pos) -join ' ')
    NewShapes; SelectOrder @('A', 'B', 'C'); $c = (Box ($slide.Shapes.Item('A'))).C
    Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' 'D V'
    Record 'DV1' 'D V: one column on the first selected, even gaps, outer edges kept' ((LineOk @('A', 'B', 'C') 'V' $c 50 $null) -and (Near (Box ($slide.Shapes.Item('B'))).B 260)) ((Pos) -join ' ')

    # Two shapes: they line up (the old Distribute did nothing with two).
    NewShapes; $slide.Shapes.Item('C').Delete(); SelectOrder @('A', 'B'); $m = (Box ($slide.Shapes.Item('A'))).M
    Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' 'D H'
    Record 'DH2' 'D H with two shapes: same row, A stays, B keeps its left' ((LineOk @('A', 'B') 'H' $m 100 $null) -and (Near $slide.Shapes.Item('B').Left 300) -and (Near $slide.Shapes.Item('A').Top 100)) ("A=$($slide.Shapes.Item('A').Left),$($slide.Shapes.Item('A').Top) B=$($slide.Shapes.Item('B').Left),$($slide.Shapes.Item('B').Top)")

    # Stacked copies (overlapping): they spread out with an 8 pt gap, a column becomes a row and a row a column.
    Stack 'H'; SelectOrder @('A', 'B', 'C')
    Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' 'D H'
    Record 'DH3' 'D H on a stacked column: a row with 8 pt gaps, in top-to-bottom order' ((LineOk @('A', 'B', 'C') 'H' 270 405 8) -and ($slide.Shapes.Item('A').Left -lt $slide.Shapes.Item('B').Left) -and ($slide.Shapes.Item('B').Left -lt $slide.Shapes.Item('C').Left)) ((Pos) -join ' ')
    Stack 'V'; SelectOrder @('A', 'B', 'C')
    Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' 'D V'
    Record 'DV3' 'D V on a stacked row: a column with 8 pt gaps, in left-to-right order' ((LineOk @('A', 'B', 'C') 'V' 375 200 8) -and ($slide.Shapes.Item('A').Top -lt $slide.Shapes.Item('B').Top) -and ($slide.Shapes.Item('B').Top -lt $slide.Shapes.Item('C').Top)) ((Pos) -join ' ')

    # The ribbon buttons run the same code as the shortcut.
    NewShapes; SelectOrder @('A', 'B', 'C'); $m = (Box ($slide.Shapes.Item('A'))).M
    Run 'ModuleObjectsAlignAndDistribute.ObjectsDistributeHorizontally'
    Record 'DR1' 'ribbon Distribute Horizontally = D H' ((LineOk @('A', 'B', 'C') 'H' $m 100 $null) -and (Near (Box ($slide.Shapes.Item('C'))).R 580)) ((Pos) -join ' ')

    # Bounds reference: the row is on the middle of the selection bounds (50..260 -> 155).
    Run 'ModulePPTKAlign.PPTK_SetAlignReference' 'Bounds'
    NewShapes; SelectOrder @('A', 'B', 'C')
    Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' 'D H'
    Record 'DB1' 'D H in bounds mode: the row is on the middle of the selection' (LineOk @('A', 'B', 'C') 'H' 155 100 $null) ((Pos) -join ' ')
    Run 'ModulePPTKAlign.PPTK_SetAlignReference' 'First'

    # One shape: centered on the slide (unchanged); one Undo restores everything.
    NewShapes; SelectOrder @('B')
    Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' 'D H'
    Record 'DS1' 'D H with one shape: centered on the slide' (Near ((Box ($slide.Shapes.Item('B'))).C) ($pres.PageSetup.SlideWidth / 2)) ((Pos) -join ' ')
    NewShapes; SelectOrder @('A', 'B', 'C'); $before = (Pos) -join ' '
    Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' 'D V'
    $after = (Pos) -join ' '
    $app.CommandBars.ExecuteMso('Undo')
    Record 'DU1' 'D V: one Undo restores all shapes' ($after -ne $before -and ((Pos) -join ' ') -eq $before) "after: $after | undone: $((Pos) -join ' ')"
} catch {
    Record 'X' 'unexpected error' $false $_.Exception.Message
} finally {
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    try { if ($addin) { $addin.Close() } } catch { }
    # Hidden before Quit: if another process still holds this instance, no empty window is left on screen.
    try { $app.Visible = 0 } catch { }
    $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    if ($null -eq $regBefore) { Remove-ItemProperty -Path $regPath -Name Reference -ErrorAction SilentlyContinue; if ((Test-Path $regPath) -and -not (Get-Item $regPath).Property.Count) { Remove-Item $regPath -ErrorAction SilentlyContinue } }
    else { Set-ItemProperty -Path $regPath -Name Reference -Value $regBefore }
}
$failed = @($results | Where-Object { -not $_ }).Count
Write-Host "TOTAL $($results.Count) tests, $failed failed"
exit $(if ($failed) { 1 } else { 0 })
