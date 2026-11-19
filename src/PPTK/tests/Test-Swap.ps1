<#
.SYNOPSIS
  Automated PowerPoint test for PPTK Swap Positions (the / shortcut, ModuleObjectsSizeAndPosition.ObjectsSwapPosition).

.DESCRIPTION
  Each shape must move to the exact centre of the other one's original position, whatever their sizes or
  rotation; sizes, rotation, text, fill, stacking order and the selection stay; two swaps give the original
  layout back; one Undo restores it. Same private-PowerPoint rules as Test-Align.ps1 (the PPTM given by -Pptm
  opened without a window, macros called through Application.Run). Exit code 0 when every case passes.
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
    function Swap { [void][System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, [object[]]@(($prefix + 'ModulePPTKShortcuts.PPTK_DispatchShortcut'), '/')) }
    $pres = $app.Presentations.Add(-1)
    $app.ActiveWindow.ViewType = 9
    function NewSlide { $s = $pres.Slides.Add($pres.Slides.Count + 1, 12); $app.ActiveWindow.View.GotoSlide($s.SlideIndex); $s }
    function Centre($x) { [pscustomobject]@{ X = $x.Left + $x.Width / 2; Y = $x.Top + $x.Height / 2 } }
    # Visible (rotated) box centre, measured independently of the frame.
    function VisCentre($x) {
        $r = $x.Rotation * [Math]::PI / 180; $w = [Math]::Abs($x.Width * [Math]::Cos($r)) + [Math]::Abs($x.Height * [Math]::Sin($r)); $h = [Math]::Abs($x.Width * [Math]::Sin($r)) + [Math]::Abs($x.Height * [Math]::Cos($r))
        $l = $x.Left + $x.Width / 2 - $w / 2; $t = $x.Top + $x.Height / 2 - $h / 2
        [pscustomobject]@{ X = $l + $w / 2; Y = $t + $h / 2 }
    }
    function Look($x) { "$([Math]::Round($x.Width,2))x$([Math]::Round($x.Height,2)) rot$($x.Rotation) fill$($x.Fill.ForeColor.RGB) z$($x.ZOrderPosition) '$($x.TextFrame2.TextRange.Text)'" }
    function Pos($s) { @($s.Shapes | ForEach-Object { "$($_.Name)@$([Math]::Round($_.Left,2)),$([Math]::Round($_.Top,2))" }) -join ' ' }
    function Case([string]$id, [string]$label, $a, $b, $s) {
        $ca = VisCentre $a; $cb = VisCentre $b; $la = Look $a; $lb = Look $b
        $s.Shapes.Range([object[]]@($a.Name, $b.Name)).Select()
        Swap
        $na = VisCentre $a; $nb = VisCentre $b
        $sel = @(1..$app.ActiveWindow.Selection.ShapeRange.Count | ForEach-Object { $app.ActiveWindow.Selection.ShapeRange.Item($_).Name }) -join ','
        $ok = (Near $na.X $cb.X) -and (Near $na.Y $cb.Y) -and (Near $nb.X $ca.X) -and (Near $nb.Y $ca.Y) -and (Look $a) -eq $la -and (Look $b) -eq $lb -and $sel -eq "$($a.Name),$($b.Name)"
        Record $id $label $ok ("A centre $([Math]::Round($ca.X,1)),$([Math]::Round($ca.Y,1)) -> $([Math]::Round($na.X,1)),$([Math]::Round($na.Y,1)); B -> $([Math]::Round($nb.X,1)),$([Math]::Round($nb.Y,1)); selection $sel")
    }

    $s = NewSlide
    $a = $s.Shapes.AddShape(1, 100, 100, 80, 60); $a.Name = 'A'; $b = $s.Shapes.AddShape(1, 300, 200, 80, 60); $b.Name = 'B'
    Case 'S1' 'same size: exact swap' $a $b $s
    Record 'S1b' 'same size: left/top swapped exactly' ((Near $a.Left 300) -and (Near $a.Top 200) -and (Near $b.Left 100) -and (Near $b.Top 100)) (Pos $s)

    $s = NewSlide
    $a = $s.Shapes.AddShape(1, 100, 100, 240, 120); $a.Name = 'Big'; $a.TextFrame2.TextRange.Text = 'Big'; $a.Fill.ForeColor.RGB = 255
    $b = $s.Shapes.AddShape(1, 500, 260, 50, 30); $b.Name = 'Small'
    Case 'S2' 'very different sizes: each centre goes to the other centre; sizes, text and fill kept' $a $b $s

    $s = NewSlide
    $a = $s.Shapes.AddShape(1, 80, 120, 300, 160); $a.Name = 'Rect'
    $b = $s.Shapes.AddShape(9, 520, 60, 40, 220); $b.Name = 'Oval'
    Case 'S3' 'large rectangle and a narrow vertical oval (the reported case)' $a $b $s

    $s = NewSlide
    $a = $s.Shapes.AddShape(1, 100, 100, 200, 60); $a.Name = 'R30'; $a.Rotation = 30
    $b = $s.Shapes.AddShape(9, 450, 250, 50, 150); $b.Name = 'O75'; $b.Rotation = 75
    Case 'S4' 'rotated shapes: visible centres swap, rotation kept' $a $b $s

    $s = NewSlide
    $a = $s.Shapes.AddShape(1, 90, 110, 220, 90); $a.Name = 'P'; $b = $s.Shapes.AddShape(9, 480, 230, 40, 160); $b.Name = 'Q'; $b.Rotation = 20
    $before = Pos $s
    $s.Shapes.Range([object[]]@('P', 'Q')).Select(); Swap; Swap
    Record 'S5' 'two swaps give the original positions back' ((Pos $s) -eq $before) "before $before | after $(Pos $s)"
    $za = $a.ZOrderPosition; $zb = $b.ZOrderPosition
    Swap
    Record 'S6' 'stacking order of each shape is kept' ($a.ZOrderPosition -eq $za -and $b.ZOrderPosition -eq $zb) "z $za,$zb -> $($a.ZOrderPosition),$($b.ZOrderPosition)"
    $moved = Pos $s
    $app.CommandBars.ExecuteMso('Undo'); Start-Sleep -Milliseconds 400
    Record 'U1' 'one Undo restores both shapes' ($moved -ne $before -and (Pos $s) -eq $before) "after swap $moved | undone $(Pos $s)"

    $s = NewSlide
    $g1 = $s.Shapes.AddShape(1, 100, 100, 160, 80); $g1.Name = 'G1'; $g2 = $s.Shapes.AddShape(9, 400, 220, 40, 120); $g2.Name = 'G2'
    $g3 = $s.Shapes.AddShape(1, 250, 400, 60, 40); $g3.Name = 'G3'
    $grp = $s.Shapes.Range([object[]]@('G1', 'G2', 'G3')).Group(); $grp.Name = 'Grp'
    $c1 = Centre $grp.GroupItems.Item('G1'); $c2 = Centre $grp.GroupItems.Item('G2'); $c3 = Centre $grp.GroupItems.Item('G3')
    $grp.GroupItems.Item('G1').Select(-1); $grp.GroupItems.Item('G2').Select(0)
    Swap
    $n1 = Centre $grp.GroupItems.Item('G1'); $n2 = Centre $grp.GroupItems.Item('G2'); $n3 = Centre $grp.GroupItems.Item('G3')
    Record 'G1' 'two shapes inside a group swap centres; the third one does not move' ((Near $n1.X $c2.X) -and (Near $n1.Y $c2.Y) -and (Near $n2.X $c1.X) -and (Near $n2.Y $c1.Y) -and (Near $n3.X $c3.X) -and (Near $n3.Y $c3.Y)) ''
} catch {
    Record 'X' 'unexpected error' $false ($_.Exception.Message + ' at line ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    try { if ($addin) { $addin.Close() } } catch { }
    try { $app.Visible = 0 } catch { }
    $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
$failed = @($results | Where-Object { -not $_ }).Count
Write-Host "TOTAL $($results.Count) tests, $failed failed"
exit $(if ($failed) { 1 } else { 0 })
