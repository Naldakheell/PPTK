<#
.SYNOPSIS
  Automated PowerPoint test for PPTK Select Similar (ModulePPTKSelectSimilar).

.DESCRIPTION
  Same rules as Test-Align.ps1: a private PowerPoint, a new unsaved presentation, the PPTM given by -Pptm opened
  without a window, macros called by qualified name through Application.Run (InvokeMember). Cases that would
  show a message box (no match, nothing selected) are not run here. Exit code 0 when every case passes.
#>
param([string]$Pptm = '')

$ErrorActionPreference = 'Stop'
if (-not $Pptm) { $Pptm = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path 'PPTK-WINDOWS-FINAL.pptm' }
$Pptm = (Resolve-Path -LiteralPath $Pptm).Path
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Write-Host 'SKIPPED: close PowerPoint first'; exit 1 }
$results = New-Object System.Collections.ArrayList
function Record([string]$id, [string]$name, [bool]$ok, [string]$detail) { [void]$results.Add($ok); Write-Host ("{0} | {1} {2} | {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $id, $name, $detail) }
function Rgb([int]$r, [int]$g, [int]$b) { $r + 256 * $g + 65536 * $b }

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
    $accent1 = $pres.SlideMaster.Theme.ThemeColorScheme.Colors(5).RGB   # msoThemeAccent1

    function Box([string]$name, [int]$type, [double]$l, [double]$t, [double]$w, [double]$h) { $x = $s.Shapes.AddShape($type, $l, $t, $w, $h); $x.Name = $name; $x }
    $r1 = Box 'R1' 1 20 20 100 60;  $r1.Fill.ForeColor.RGB = (Rgb 255 0 0); $r1.Line.ForeColor.RGB = (Rgb 0 0 255)
    $r2 = Box 'R2' 1 140 20 80 50;  $r2.Fill.ForeColor.RGB = (Rgb 255 0 0); $r2.Line.ForeColor.RGB = (Rgb 0 0 255)
    $r3 = Box 'R3' 9 240 20 100 60; $r3.Fill.ForeColor.RGB = (Rgb 255 0 0); $r3.Line.ForeColor.RGB = (Rgb 0 128 0)   # oval
    $r4 = Box 'R4' 1 360 20 100 60; $r4.Fill.ForeColor.ObjectThemeColor = 5; $r4.Line.Visible = 0                   # theme Accent 1
    $r5 = Box 'R5' 1 480 20 100 60; $r5.Fill.ForeColor.RGB = $accent1; $r5.Line.Visible = 0                          # same color as explicit RGB
    $hid = Box 'Hidden' 1 600 20 100 60; $hid.Fill.ForeColor.RGB = (Rgb 255 0 0); $hid.Visible = 0
    $w = Box 'W' 9 600 120 80 200; $w.Fill.Visible = 0; $w.Line.Visible = 0                                         # oval, as wide as R2

    function Text([string]$name, [double]$t, [string]$text) { $x = $s.Shapes.AddTextbox(1, 20, $t, 300, 30); $x.Name = $name; $x.TextFrame2.TextRange.Text = $text; $x }
    $t1 = Text 'T1' 120 'Hello'; $t1.TextFrame2.TextRange.Font.Name = 'Arial'; $t1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (Rgb 0 0 0)
    $t2 = Text 'T2' 160 'one two'; $t2.TextFrame2.TextRange.Font.Name = 'Calibri'; $t2.TextFrame2.TextRange.Characters(5, 3).Font.Name = 'Arial'
    $t2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (Rgb 255 0 0); $t2.TextFrame2.TextRange.Characters(5, 3).Font.Fill.ForeColor.RGB = (Rgb 0 0 0)
    $arabic = [string]::new([char[]](0x0645, 0x0631, 0x062D, 0x0628, 0x0627))
    $t3 = Text 'T3' 200 $arabic; $t3.TextFrame2.TextRange.Font.Name = 'Arial'; $t3.TextFrame2.TextRange.Font.NameComplexScript = 'Tahoma'; $t3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (Rgb 0 0 255)
    $t4 = Text 'T4' 240 $arabic; $t4.TextFrame2.TextRange.Font.Name = 'Calibri'; $t4.TextFrame2.TextRange.Font.NameComplexScript = 'Tahoma'
    $g1 = Text 'G1' 280 'grouped'; $g1.TextFrame2.TextRange.Font.Name = 'Arial'
    $g2 = Text 'G2' 320 'other'; $g2.TextFrame2.TextRange.Font.Name = 'Calibri'
    $grp = $s.Shapes.Range([object[]]@('G1', 'G2')).Group(); $grp.Name = 'Group'
    $tbl = $s.Shapes.AddTable(1, 2, 400, 300, 200, 40); $tbl.Name = 'Table'
    $c = $tbl.Table.Cell(1, 2).Shape.TextFrame2.TextRange; $c.Text = 'cell'; $c.Font.Name = 'Arial'
    $c.Font.Fill.ForeColor.RGB = (Rgb 0 0 0)   # the default table style makes header text white

    function Pick([string]$name) { $s.Shapes.Item($name).Select(-1) }
    function Selected { $sr = $app.ActiveWindow.Selection.ShapeRange; $n = @(); for ($i = 1; $i -le $sr.Count; $i++) { $n += $sr.Item($i).Name }; $n }
    function Check([string]$id, [string]$label, [string]$ref, [string]$token, [string[]]$expect) {
        Pick $ref
        Run 'ModulePPTKShortcuts.PPTK_DispatchShortcut' $token | Out-Null
        $got = @(Selected)
        $ok = ($got.Count -gt 0 -and $got[0] -eq $ref -and ((@($got | Sort-Object) -join ',') -eq (@($expect | Sort-Object) -join ',')))
        Record $id $label $ok ("got " + ($got -join ',') + " | expected " + ($expect -join ','))
    }

    Check 'F1' 'same fill: red shapes, hidden one excluded, reference first' 'R1' 'CMD SELECT FILL' @('R1', 'R2', 'R3')
    Check 'F2' 'same fill: theme Accent 1 matches the same explicit RGB' 'R4' 'CMD SELECT FILL' @('R4', 'R5')
    Check 'O1' 'same outline color' 'R1' 'CMD SELECT OUTLINE' @('R1', 'R2')
    Check 'Y1' 'same shape type: rectangles, not the oval or text boxes' 'R1' 'CMD SELECT TYPE' @('R1', 'R2', 'R4', 'R5')
    Check 'Z1' 'same size' 'R1' 'CMD SELECT SIZE' @('R1', 'R3', 'R4', 'R5')
    Check 'N1' 'same font: mixed text, group and table text count; Arabic-only text does not' 'T1' 'CMD SELECT FONT' @('T1', 'T2', 'Group', 'Table')
    Check 'N2' 'same Arabic font (complex script)' 'T3' 'CMD SELECT FONT' @('T3', 'T4')
    Check 'C1' 'same font color: any run with that color' 'T1' 'CMD SELECT FONT COLOR' @('T1', 'T2', 'Group', 'Table', 'T4')

    # New "same as first selected" commands
    $t1.TextFrame2.TextRange.Font.Size = 28; $t2.TextFrame2.TextRange.Characters(5, 3).Font.Size = 28
    Check 'S1' 'same font size: any run with that size' 'T1' 'CMD SELECT FONT SIZE' @('T1', 'T2')
    Check 'W1' 'same width' 'R2' 'CMD SELECT WIDTH' @('R2', 'W')
    Check 'H1' 'same height' 'R1' 'CMD SELECT HEIGHT' @('R1', 'R3', 'R4', 'R5')

    # Value lists for the right-click Select menu (values on the current slide, nothing selected first)
    $app.ActiveWindow.Selection.Unselect()
    function Values([string]$kind) { $n = [int](Run 'ModulePPTKSelectSimilar.PPTK_SelectCollect' $kind); @(for ($i = 1; $i -le $n; $i++) { [pscustomobject]@{ Value = [string](Run 'ModulePPTKSelectSimilar.PPTK_SelectValueAt' "$i"); Label = [string](Run 'ModulePPTKSelectSimilar.PPTK_SelectLabelAt' "$i") } }) }
    function ByValue([string]$id, [string]$label, [string]$spec, [string[]]$expect) {
        $app.ActiveWindow.Selection.Unselect()
        $n = [int](Run 'ModulePPTKSelectSimilar.PPTK_SelectByValue' $spec)
        $got = @(Selected)
        Record $id $label ($n -eq $expect.Count -and ((@($got | Sort-Object) -join ',') -eq (@($expect | Sort-Object) -join ','))) ("returned $n; got " + ($got -join ',') + " | expected " + ($expect -join ','))
    }
    $fill = Values 'FILL'
    Record 'V1' 'fill values: red and Accent 1 once each, hidden and no-fill not listed' ((@($fill | Where-Object Value -eq "$(Rgb 255 0 0)").Count -eq 1) -and (@($fill | Where-Object Value -eq "$accent1").Count -eq 1) -and (@($fill | Where-Object Label -eq '#FF0000').Count -eq 1)) (($fill | ForEach-Object { "$($_.Value)=$($_.Label)" }) -join ' ; ')
    $types = Values 'TYPE'
    $labels = @($types | ForEach-Object Label)
    Record 'V2' 'object type values have names' ((@('Rectangle', 'Oval', 'Text Box', 'Group', 'Table') | Where-Object { $labels -notcontains $_ }).Count -eq 0) ($labels -join ' ; ')
    $sizes = Values 'FONT SIZE'
    Record 'V3' 'font size values: 18 pt and 28 pt' ((@($sizes | ForEach-Object Label) -contains '28 pt') -and (@($sizes | ForEach-Object Label) -contains '18 pt')) ((@($sizes | ForEach-Object Label)) -join ' ; ')
    $widths = Values 'WIDTH'
    Record 'V4' 'width values: 80.0 pt listed once' (@($widths | Where-Object Value -eq '80.0').Count -eq 1) ((@($widths | ForEach-Object Label)) -join ' ; ')
    ByValue 'P1' 'pick red fill: R1, R2, R3 (hidden one excluded)' "FILL|$(Rgb 255 0 0)" @('R1', 'R2', 'R3')
    $oval = ($types | Where-Object Label -eq 'Oval' | Select-Object -First 1).Value
    ByValue 'P2' 'pick type Oval: R3 and W' "TYPE|$oval" @('R3', 'W')
    ByValue 'P3' 'pick font size 28: T1 and T2' 'FONT SIZE|28' @('T1', 'T2')
    ByValue 'P4' 'pick width 80.0: R2 and W' 'WIDTH|80.0' @('R2', 'W')
    ByValue 'P5' 'pick size 100.0x60.0' 'SIZE|100.0x60.0' @('R1', 'R3', 'R4', 'R5')
    $app.ActiveWindow.Selection.Unselect()
    $n = [int](Run 'ModulePPTKSelectSimilar.PPTK_SelectAllText')
    $got = @(Selected)
    $expect = @('T1', 'T2', 'T3', 'T4', 'Group', 'Table')
    Record 'A1' 'Select All Text on Slide: text boxes, group with text, table' ($n -eq 6 -and ((@($got | Sort-Object) -join ',') -eq (@($expect | Sort-Object) -join ','))) ("returned $n; got " + ($got -join ','))

    # Submenu XML (getContent): valid, swatch gallery without text for colors, value buttons otherwise
    foreach ($kind in 'FILL', 'OUTLINE', 'FONT COLOR', 'FONT SIZE', 'TYPE', 'WIDTH', 'HEIGHT', 'SIZE') {
        $xml = [string](Run 'ModulePPTKSelectMenu.PPTK_SelectMenuXml' $kind); $ok = $true; $detail = ''
        try { [xml]$d = $xml } catch { $ok = $false; $detail = 'invalid XML: ' + $xml }
        if ($ok) {
            if ($kind -in 'FILL', 'OUTLINE', 'FONT COLOR') {
                $g = $d.menu.gallery
                $ok = ($null -ne $g) -and $g.showItemLabel -eq 'false' -and $g.tag -eq $kind -and ($null -eq $d.menu.button)
                $detail = "gallery $($g.id), tag $($g.tag)"
            } else {
                $b = @($d.menu.button)
                $ok = $b.Count -gt 0 -and -not ($b | Where-Object { $_.tag -notlike "$kind|*" -or $_.onAction -ne 'PPTK_SelectMenuPick' })
                $detail = (@($b | ForEach-Object { $_.label }) -join ' ; ')
            }
        }
        Record "M-$($kind -replace ' ', '')" "Select By $kind submenu" $ok $detail
    }
    $x1 = [string](Run 'ModulePPTKSelectMenu.PPTK_SelectMenuXml' 'TYPE'); $x2 = [string](Run 'ModulePPTKSelectMenu.PPTK_SelectMenuXml' 'TYPE')
    Record 'M-ID' 'each open gets new control ids' (([xml]$x1).menu.button[0].id -ne ([xml]$x2).menu.button[0].id) ''
    $swatch = [string](Run 'ModulePPTKSelectMenu.PPTK_SwatchFile' '192')
    $okFile = $false
    if (Test-Path $swatch) { $bytes = [IO.File]::ReadAllBytes($swatch); $center = 54 + (8 * 16 + 8) * 3; $okFile = ($bytes[0] -eq 66 -and $bytes[1] -eq 77 -and $bytes[$center] -eq 0 -and $bytes[$center + 1] -eq 0 -and $bytes[$center + 2] -eq 192) }
    Record 'M-SW' 'swatch BMP has the color (#C00000)' $okFile $swatch

    $before = @($s.Shapes | ForEach-Object { "$($_.Name)=$($_.Left),$($_.Top),$($_.Width),$($_.Height)" }) -join ';'
    Pick 'R1'; Run 'ModulePPTKSelectSimilar.PPTK_SelectSimilarFill' | Out-Null
    $after = @($s.Shapes | ForEach-Object { "$($_.Name)=$($_.Left),$($_.Top),$($_.Width),$($_.Height)" }) -join ';'
    Record 'K1' 'selecting changes nothing on the slide' ($before -eq $after) ''
} catch {
    Record 'X' 'unexpected error' $false $_.Exception.Message
} finally {
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    try { if ($addin) { $addin.Close() } } catch { }
    $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
$failed = @($results | Where-Object { -not $_ }).Count
Write-Host "TOTAL $($results.Count) tests, $failed failed"
exit $(if ($failed) { 1 } else { 0 })
