<#
.SYNOPSIS
  Automated PowerPoint test for PPTK rows and columns: Ctrl+2 / Ctrl+3 / Alt+2 / Alt+3 (ModulePPTKGrid).

.DESCRIPTION
  Issue 1 (grow mode): shapes that must keep their size and multi-object designs are copied or removed whole,
  outward, with size, rotation, text, formatting, offsets and stacking order kept; partly overlapping objects
  that may or may not be one design are left alone with a message.
  Issue 2 (divide mode): one rectangle divided into rows and columns gets 8 pt gaps and keeps its outer bounds.
  Also: existing keep-bounds grids, tables, ragged selections, and one Undo per command.

  A private PowerPoint, a new unsaved presentation, the PPTM given by -Pptm opened without a window, macros
  called by qualified name (Application.Run through InvokeMember) with the same tokens the keys send. Commands
  that show a message run from a second process; the message box is read and closed through UI Automation
  (its OK button), never with keystrokes. Exit code 0 when every case passes.
#>
param([string]$Pptm = '')

$ErrorActionPreference = 'Stop'
if (-not $Pptm) { $Pptm = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path 'PPTK-WINDOWS-FINAL.pptm' }
$Pptm = (Resolve-Path -LiteralPath $Pptm).Path
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Write-Host 'SKIPPED: close PowerPoint first'; exit 1 }
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$results = New-Object System.Collections.ArrayList
function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) { [void]$results.Add($ok); Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail) }
function Near([double]$a, [double]$b, [double]$tol = 0.1) { [Math]::Abs($a - $b) -le $tol }
function R2([double]$v) { [Math]::Round($v, 2) }

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$addin = $null; $pres = $null
$runner = Join-Path $env:TEMP ('pptk-grid-runner-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.ps1')
try {
    $addin = $app.Presentations.Open($Pptm, -1, 0, 0)
    $prefix = "'" + $addin.Name + "'!"
    function Run([string]$macro, [string]$arg = $null) {
        $a = if ($arg) { [object[]]@(($prefix + $macro), $arg) } else { [object[]]@(($prefix + $macro)) }
        return [System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, $a)
    }
    function Key([string]$token) { Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' $token | Out-Null }
    $pres = $app.Presentations.Add(-1)
    $app.ActiveWindow.ViewType = 9
    $pp = (Get-Process POWERPNT | Sort-Object StartTime | Select-Object -Last 1).Id

    # A command that is expected to show a PPTK message: run it from a second process, read and close the box.
    @'
param([string]$macro, [string]$token)
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
[void][System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, [object[]]@($macro, $token))
'@ | Set-Content $runner -Encoding UTF8
    function KeyWithMessage([string]$token) {
        # Quoted: Start-Process joins arguments with spaces, so "CTRL 2" would arrive as two arguments.
        $p = Start-Process powershell -ArgumentList '-NoProfile', '-File', $runner, ('"' + $prefix + 'ModulePPTKShortcuts.PPTK_DispatchShortcut"'), ('"' + $token + '"') -PassThru -WindowStyle Hidden
        $text = $null
        $root = [Windows.Automation.AutomationElement]::RootElement
        $pc = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $pp)
        $named = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, 'PPTK')
        for ($i = 0; $i -lt 60 -and -not $text; $i++) {
            Start-Sleep -Milliseconds 250
            # The VBA message box is a child of the PowerPoint window; its text and OK button are panes.
            $box = $null
            foreach ($frame in @($root.FindAll([Windows.Automation.TreeScope]::Children, $pc))) {
                $box = $frame.FindFirst([Windows.Automation.TreeScope]::Children, $named)
                if ($box) { break }
            }
            if ($box) {
                $parts = @($box.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition))
                $text = (@($parts | Where-Object { $_.Current.Name -and $_.Current.Name -ne 'OK' } | ForEach-Object { $_.Current.Name }) -join ' ')
                if (-not $text) { $text = '(no text)' }
                $closed = $false
                $ok = @($parts | Where-Object { $_.Current.Name -eq 'OK' }) | Select-Object -First 1
                if ($ok) { try { ($ok.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern)).Invoke(); $closed = $true } catch { } }
                if (-not $closed) { ($box.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)).Close() }
            }
        }
        if (-not $p.WaitForExit(20000)) { $p.Kill() }
        return $text
    }

    function NewSlide { $s = $pres.Slides.Add($pres.Slides.Count + 1, 12); $app.ActiveWindow.View.GotoSlide($s.SlideIndex); $s }
    function Add($s, [int]$type, [string]$name, [double]$l, [double]$t, [double]$w, [double]$h) { $x = $s.Shapes.AddShape($type, $l, $t, $w, $h); $x.Name = $name; $x }
    function Pick($s, [string[]]$names) { $s.Shapes.Range([object[]]$names).Select() }
    function Snap($x) {
        $txt = ''; try { if ($x.HasTextFrame) { $txt = $x.TextFrame2.TextRange.Text } } catch { }
        "$(R2 $x.Width)x$(R2 $x.Height) rot$(R2 $x.Rotation) fill$($x.Fill.ForeColor.RGB) line$($x.Line.ForeColor.RGB) '$txt'"
    }
    function State($s) { @(for ($i = 1; $i -le $s.Shapes.Count; $i++) { $x = $s.Shapes.Item($i); "$($x.Name)@$(R2 $x.Left),$(R2 $x.Top) $(Snap $x)" }) -join ' | ' }
    # The shapes added after the first $n (copies), in stacking order.
    function Added($s, [int]$n) { @(for ($i = $n + 1; $i -le $s.Shapes.Count; $i++) { $s.Shapes.Item($i) }) }
    function VisBox($x) {
        $rad = $x.Rotation * [Math]::PI / 180; $w = [Math]::Abs($x.Width * [Math]::Cos($rad)) + [Math]::Abs($x.Height * [Math]::Sin($rad)); $h = [Math]::Abs($x.Width * [Math]::Sin($rad)) + [Math]::Abs($x.Height * [Math]::Cos($rad))
        $cx = $x.Left + $x.Width / 2; $cy = $x.Top + $x.Height / 2
        [pscustomobject]@{ L = $cx - $w / 2; T = $cy - $h / 2; R = $cx + $w / 2; B = $cy + $h / 2 }
    }
    function Undo { $app.CommandBars.ExecuteMso('Undo') }

    # A composite design: ring, oval inside, text box inside the oval, small star on the oval. Red, blue, text "Hi".
    function Design($s, [string]$p, [double]$x, [double]$y) {
        $ring = Add $s 18 "$p-ring" $x $y 140 140; $ring.Fill.ForeColor.RGB = 255
        $oval = Add $s 9 "$p-oval" ($x + 20) ($y + 20) 100 100; $oval.Fill.ForeColor.RGB = 16711680
        $tb = $s.Shapes.AddTextbox(1, $x + 30, $y + 55, 80, 30); $tb.Name = "$p-text"; $tb.TextFrame2.AutoSize = 0; $tb.Height = 30; $tb.TextFrame2.TextRange.Text = 'Hi'
        $star = Add $s 92 "$p-star" ($x + 30) ($y + 30) 20 20; $star.Rotation = 15
        @("$p-ring", "$p-oval", "$p-text", "$p-star")
    }

    # ================================================================ Issue 1: grow mode
    $s = NewSlide; $o = Add $s 9 'Oval' 100 100 100 60; $o.TextFrame2.TextRange.Text = 'A'; $before = Snap $o
    Pick $s @('Oval'); Key 'CTRL 2'
    $c = @(Added $s 1)[0]
    Record 'G1' 'oval Ctrl+2: copy beside it, same size and text, original untouched' ($s.Shapes.Count -eq 2 -and (Snap $c) -eq $before -and (Snap $o) -eq $before -and (Near $c.Left 208) -and (Near $c.Top 100) -and (Near $o.Left 100)) (State $s)

    $s = NewSlide; $o = Add $s 9 'Oval' 100 100 100 60; Pick $s @('Oval'); Key 'CTRL 3'
    $c = @(Added $s 1)[0]
    Record 'G2' 'oval Ctrl+3: copy below, same size' ($s.Shapes.Count -eq 2 -and (Near $c.Top 168) -and (Near $c.Left 100) -and (Near $c.Width 100) -and (Near $c.Height 60)) (State $s)

    $s = NewSlide; $a = Add $s 9 'O1' 100 100 100 50; $b = Add $s 9 'O2' 260 100 100 50; $a.Rotation = 30; $b.Rotation = 30
    $gapBefore = (VisBox $b).L - (VisBox $a).R
    Pick $s @('O1', 'O2'); Key 'CTRL 2'
    $c = @(Added $s 2)[0]
    Record 'G3' 'rotated ovals: rotation and size kept, same visible gap' ($s.Shapes.Count -eq 3 -and (Near $c.Rotation 30) -and (Near $c.Width 100) -and (Near $c.Height 50) -and (Near ((VisBox $c).L - (VisBox $b).R) $gapBefore) -and (Near $c.Top $b.Top)) ("gap $(R2 $gapBefore) | " + (State $s))

    $s = NewSlide; $a = Add $s 9 'Small' 100 100 60 60; $b = Add $s 9 'Wide' 175 110 100 40
    Pick $s @('Small', 'Wide'); Key 'CTRL 2'
    $c = @(Added $s 2)[0]
    Record 'G4' 'different sizes: the edge shape is copied at its own size, measured gap 15' ($s.Shapes.Count -eq 3 -and (Near $c.Width 100) -and (Near $c.Height 40) -and (Near $c.Left 290) -and (Near $c.Top 110) -and (Near $a.Width 60)) (State $s)

    $s = NewSlide; $names = Design $s 'D' 100 100
    $orig = @($names | ForEach-Object { $s.Shapes.Item($_) })
    Pick $s $names; Key 'CTRL 2'
    $copies = @(Added $s 4)
    $dx = $copies[0].Left - $orig[0].Left
    $okMembers = $copies.Count -eq 4 -and (Near $dx 148)
    for ($i = 0; $i -lt 4 -and $okMembers; $i++) {
        if ((Snap $copies[$i]) -ne (Snap $orig[$i]) -or -not (Near ($copies[$i].Left - $orig[$i].Left) $dx) -or -not (Near $copies[$i].Top $orig[$i].Top)) { $okMembers = $false }
    }
    Record 'G5' 'design (ring, oval, text, rotated star): copied whole; sizes, rotation, text, colors, offsets and stacking order kept' $okMembers ("dx $(R2 $dx) | " + (State $s))
    $origState = @($orig | ForEach-Object { "$(R2 $_.Left),$(R2 $_.Top) $(Snap $_)" }) -join ';'
    Undo
    Record 'U1' 'one Undo removes the whole copied design' ($s.Shapes.Count -eq 4 -and ((@($names | ForEach-Object { $x = $s.Shapes.Item($_); "$(R2 $x.Left),$(R2 $x.Top) $(Snap $x)" }) -join ';') -eq $origState)) (State $s)

    $s = NewSlide; $n1 = Design $s 'D' 60 100; $n2 = Design $s 'E' 230 100
    Pick $s ($n1 + $n2); Key 'CTRL 2'
    $copies = @(Added $s 8)
    Record 'G6' 'two designs in a row: the right design is copied at the measured gap (30)' ($copies.Count -eq 4 -and (Near ($copies[0].Left - $s.Shapes.Item('E-ring').Left) 170)) (State $s)
    $s.Shapes.Range().Select()   # the copies share their originals' names, so select every shape
    Key 'ALT 2'
    Record 'G7' 'Alt+2 removes the last design (all 4 objects); the others do not move' ($s.Shapes.Count -eq 8 -and (Near $s.Shapes.Item('E-ring').Left 230) -and (Near $s.Shapes.Item('D-ring').Left 60)) (State $s)
    Undo
    Record 'U2' 'one Undo after Alt+2 brings the deleted design back' ($s.Shapes.Count -eq 12) (State $s)

    $s = NewSlide; $names = Design $s 'D' 100 100; $g = $s.Shapes.Range([object[]]$names).Group(); $g.Name = 'G'
    $gw = $g.Width; $gh = $g.Height
    Pick $s @('G'); Key 'CTRL 2'
    $c = @(Added $s 1)[0]
    Record 'G8' 'grouped design: copied at its size with its 4 objects' ($s.Shapes.Count -eq 2 -and $c.Type -eq 6 -and $c.GroupItems.Count -eq 4 -and (Near $c.Width $gw) -and (Near $c.Height $gh) -and (Near $c.Left ($g.Left + $gw + 8))) (State $s)

    $s = NewSlide; $a = Add $s 9 'V1' 100 100 100 100; $b = Add $s 9 'V2' 170 100 100 100; $before = State $s
    Pick $s @('V1', 'V2'); $msg = KeyWithMessage 'CTRL 2'
    Record 'A1' 'Venn circles (30% overlap): ambiguous, message, nothing changes' ($s.Shapes.Count -eq 2 -and (State $s) -eq $before -and $msg -match 'overlap partly') "message: $msg"

    $s = NewSlide; $a = Add $s 9 'V1' 100 100 100 100; $b = Add $s 9 'V2' 195 100 100 100
    Pick $s @('V1', 'V2'); Key 'CTRL 2'
    Record 'A2' 'circles touching by 5% are separate shapes (not a design), copied' ($s.Shapes.Count -eq 3) (State $s)

    $s = NewSlide
    foreach ($k in 0, 1, 2) {
        $card = Add $s 5 "Card$k" (100 + $k * 170) 100 150 100
        $tb = $s.Shapes.AddTextbox(1, 110 + $k * 170, 130, 130, 30); $tb.Name = "Label$k"; $tb.TextFrame2.AutoSize = 0; $tb.Height = 30; $tb.TextFrame2.TextRange.Text = "Card $k"
    }
    $s.Shapes.Range().Select(); Key 'CTRL 2'
    $copies = @(Added $s 6)
    Record 'C1' 'cards with a label inside: the last card and its label are copied together, text kept, cards not resized' ($copies.Count -eq 2 -and (Near $copies[0].Left 610) -and (Near $copies[0].Width 150) -and $copies[1].TextFrame2.TextRange.Text -eq 'Card 2' -and (Near ($copies[1].Left - $copies[0].Left) 10) -and (Near $s.Shapes.Item('Card0').Width 150)) (State $s)

    $s = NewSlide; $card = Add $s 1 'Card' 100 100 150 100; $tb = $s.Shapes.AddTextbox(1, 200, 120, 100, 30); $tb.Name = 'Half'; $tb.TextFrame2.AutoSize = 0; $tb.Height = 30; $tb.TextFrame2.TextRange.Text = 'Half out'
    $before = State $s
    $s.Shapes.Range().Select(); $msg = KeyWithMessage 'CTRL 2'
    Record 'A3' 'a label half outside its card: ambiguous, message, nothing changes' ((State $s) -eq $before -and $msg -match 'overlap partly') "message: $msg"

    $s = NewSlide
    foreach ($r in 0, 1) { foreach ($k in 0, 1, 2) { $x = Add $s 9 "C$r$k" (100 + $k * 120) (100 + $r * 90) 100 60 } }
    Pick $s @('C00', 'C01', 'C02', 'C10', 'C11', 'C12'); Key 'ALT 3'
    Record 'G9' 'Alt+3 on two rows of ovals removes the bottom row, nothing resized' ($s.Shapes.Count -eq 3 -and (Near $s.Shapes.Item('C00').Width 100) -and (Near $s.Shapes.Item('C02').Left 340)) (State $s)
    Pick $s @('C00', 'C01', 'C02'); Key 'CTRL 3'
    $copies = @(Added $s 3)
    Record 'G10' 'Ctrl+3 on a row of ovals: a new row below at the other axis gap (20)' ($copies.Count -eq 3 -and (Near $copies[0].Top 180) -and (Near $copies[2].Left 340)) (State $s)

    # ================================================================ Issue 2: divide one rectangle
    function Cells($s) { @(for ($i = 1; $i -le $s.Shapes.Count; $i++) { $s.Shapes.Item($i) }) }
    function Gaps($cells, [bool]$vertical) {
        $out = @()
        foreach ($a in $cells) { foreach ($b in $cells) {
            if ($vertical) { if ((Near $a.Left $b.Left 1) -and $b.Top -gt $a.Top + 1) { $nearest = @($cells | Where-Object { (Near $_.Left $a.Left 1) -and $_.Top -gt $a.Top + 1 } | Sort-Object Top)[0]; if ($nearest -eq $b) { $out += $b.Top - ($a.Top + $a.Height) } } }
            else { if ((Near $a.Top $b.Top 1) -and $b.Left -gt $a.Left + 1) { $nearest = @($cells | Where-Object { (Near $_.Top $a.Top 1) -and $_.Left -gt $a.Left + 1 } | Sort-Object Left)[0]; if ($nearest -eq $b) { $out += $b.Left - ($a.Left + $a.Width) } } }
        } }
        $out
    }
    function Bounds($cells) { $l = ($cells | ForEach-Object { $_.Left } | Measure-Object -Minimum).Minimum; $t = ($cells | ForEach-Object { $_.Top } | Measure-Object -Minimum).Minimum; $r = ($cells | ForEach-Object { $_.Left + $_.Width } | Measure-Object -Maximum).Maximum; $b = ($cells | ForEach-Object { $_.Top + $_.Height } | Measure-Object -Maximum).Maximum; "$(R2 $l),$(R2 $t),$(R2 $r),$(R2 $b)" }
    function Overlaps($cells) { foreach ($a in $cells) { foreach ($b in $cells) { if ($a -ne $b -and $a.Left -lt $b.Left + $b.Width - 0.5 -and $b.Left -lt $a.Left + $a.Width - 0.5 -and $a.Top -lt $b.Top + $b.Height - 0.5 -and $b.Top -lt $a.Top + $a.Height - 0.5) { return $true } } }; $false }
    function SelectAll($s) { $s.Shapes.Range().Select() }

    $s = NewSlide; $r = Add $s 1 'Rect' 100 100 400 300; $r.TextFrame2.TextRange.Text = 'Title'
    Pick $s @('Rect'); Key 'CTRL 3'; SelectAll $s; Key 'CTRL 3'
    $cells = Cells $s; $g = Gaps $cells $true
    Record 'D1' 'one rectangle into 3 rows: gaps 8, same heights, bounds kept, no overlap' ($cells.Count -eq 3 -and $g.Count -eq 2 -and @($g | Where-Object { -not (Near $_ 8) }).Count -eq 0 -and (Bounds $cells) -eq '100,100,500,400' -and -not (Overlaps $cells) -and (Near $cells[0].Height $cells[2].Height)) ("gaps " + ($g | ForEach-Object { R2 $_ }) + " | bounds " + (Bounds $cells))
    Record 'D4' 'the original keeps its text; new cells are empty' ($s.Shapes.Item('Rect').TextFrame2.TextRange.Text -eq 'Title' -and @($cells | Where-Object { $_.Name -ne 'Rect' -and $_.TextFrame2.TextRange.Text -ne '' }).Count -eq 0) ''

    $s = NewSlide; $r = Add $s 1 'Rect' 100 100 400 300
    Pick $s @('Rect'); Key 'CTRL 2'; SelectAll $s; Key 'CTRL 2'; SelectAll $s; Key 'CTRL 2'
    $cells = Cells $s; $g = Gaps $cells $false
    Record 'D2' 'one rectangle into 4 columns: gaps 8, same widths, bounds kept' ($cells.Count -eq 4 -and $g.Count -eq 3 -and @($g | Where-Object { -not (Near $_ 8) }).Count -eq 0 -and (Bounds $cells) -eq '100,100,500,400' -and (Near $cells[0].Width $cells[3].Width)) ("gaps " + ($g | ForEach-Object { R2 $_ }) + " | widths " + ($cells | ForEach-Object { R2 $_.Width }))
    SelectAll $s; Key 'CTRL 3'; SelectAll $s; Key 'CTRL 3'
    $cells = Cells $s; $gh = Gaps $cells $false; $gv = Gaps $cells $true
    Record 'D3' 'then 3 rows: 3 x 4 grid, every gap 8 both ways, bounds kept, no overlap' ($cells.Count -eq 12 -and $gh.Count -eq 9 -and $gv.Count -eq 8 -and @(@($gh) + @($gv) | Where-Object { -not (Near $_ 8) }).Count -eq 0 -and (Bounds $cells) -eq '100,100,500,400' -and -not (Overlaps $cells)) ("h " + ($gh | ForEach-Object { R2 $_ }) + " | v " + ($gv | ForEach-Object { R2 $_ }) + " | bounds " + (Bounds $cells))

    $s = NewSlide; $r = Add $s 1 'Rect' 100 100 400 300; Pick $s @('Rect'); Key 'CTRL 2'
    Undo
    Record 'U3' 'one Undo after dividing restores the single rectangle' ($s.Shapes.Count -eq 1 -and (Near $s.Shapes.Item('Rect').Width 400)) (State $s)

    # ================================================================ existing behavior
    $s = NewSlide; foreach ($k in 0, 1, 2) { $x = Add $s 1 "R$k" (100 + $k * 120) 100 100 60 }
    SelectAll $s; Key 'CTRL 2'
    $cells = Cells $s; $g = Gaps $cells $false
    Record 'K1' 'rectangle grid with 20 pt gaps: 4 columns in the same bounds, gaps stay 20' ($cells.Count -eq 4 -and (Bounds $cells) -eq '100,100,440,160' -and @($g | Where-Object { -not (Near $_ 20) }).Count -eq 0 -and (Near $cells[0].Width 70)) ("gaps " + ($g | ForEach-Object { R2 $_ }) + " | widths " + ($cells | ForEach-Object { R2 $_.Width }))
    SelectAll $s; Key 'ALT 2'
    $cells = Cells $s; $g = Gaps $cells $false
    Record 'K2' 'Alt+2: back to 3 columns in the same bounds, gaps 20' ($cells.Count -eq 3 -and (Bounds $cells) -eq '100,100,440,160' -and @($g | Where-Object { -not (Near $_ 20) }).Count -eq 0) ("widths " + ($cells | ForEach-Object { R2 $_.Width }))

    $s = NewSlide; foreach ($k in 0, 1, 2) { $x = Add $s 1 "R$k" (100 + $k * 100) 100 100 60 }
    SelectAll $s; Key 'CTRL 2'
    $cells = Cells $s; $g = Gaps $cells $false
    Record 'K3' 'touching rectangle grid stays touching' ($cells.Count -eq 4 -and (Bounds $cells) -eq '100,100,400,160' -and @($g | Where-Object { -not (Near $_ 0) }).Count -eq 0) ("gaps " + ($g | ForEach-Object { R2 $_ }))

    $s = NewSlide; $t = $s.Shapes.AddTable(2, 2, 100, 100, 300, 80); $t.Name = 'T'; $w = $t.Width
    Pick $s @('T'); Key 'CTRL 2'
    Record 'K4' 'table: a third column, same width' ($t.Table.Columns.Count -eq 3 -and (Near $t.Width $w 1)) "width $(R2 $t.Width)"

    $s = NewSlide; foreach ($k in 0, 1, 2) { $x = Add $s 1 "R$k" (100 + $k * 120) 100 100 60 }; $x = Add $s 1 'R3' 100 200 100 60; $before = State $s
    SelectAll $s; $msg = KeyWithMessage 'CTRL 2'
    Record 'K5' 'ragged selection: message, nothing changes' ((State $s) -eq $before -and $msg -match 'one grid') "message: $msg"
} catch {
    Record 'X' 'unexpected error' $false ($_.Exception.Message + ' at line ' + $_.InvocationInfo.ScriptLineNumber + ': ' + $_.InvocationInfo.Line.Trim())
} finally {
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    try { if ($addin) { $addin.Close() } } catch { }
    try { $app.Visible = 0 } catch { }
    $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    Remove-Item $runner -ErrorAction SilentlyContinue
}
$failed = @($results | Where-Object { -not $_ }).Count
Write-Host "TOTAL $($results.Count) tests, $failed failed"
exit $(if ($failed) { 1 } else { 0 })
