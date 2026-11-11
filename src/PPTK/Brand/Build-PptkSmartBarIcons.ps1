# Builds the Smart Bar icon set: 32px 8-bit BMPs for the dark palette.
#
# Most icons reuse the PPTK ribbon artwork in icons\32. The dark ribbon ink
# (#30343B) is recolored to the Smart Bar secondary text color (#AFC1CD), cyan
# (#23BAEC) stays, and white cutouts become transparent so the row color shows
# through. Icons that had no clear ribbon artwork are drawn here on the same
# 16-unit grid as Generate-PptkRibbonIcons.ps1 (2 px per unit at 32 px).
#
# BMP is used because MSForms LoadPicture cannot read PNG, a GIF stored in a
# form loses its transparency, and a stored ICO is reduced to 16 colors. Each
# icon is composited on the normal row color (<name>.bmp) and on the selected
# row color (<name>-sel.bmp).
#
# Output: smartbar-icons\<name>.bmp, <name>-sel.bmp, and contact-sheet.png.
#
# -Ribbon writes the icons drawn here as transparent PNGs in the ribbon style
# (ink #30343B) to icons\16 and icons\32, so the ribbon shows the same
# concepts as the Smart Bar. Recolored ribbon icons are not rewritten.
# Does not edit the PPTM, forms, or the ribbon icons.
param(
    [string]$OutDir = '',
    [switch]$Ribbon
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Ribbon32 = Join-Path $Root 'icons\32'
if ($OutDir -eq '') { $OutDir = Join-Path $Root 'smartbar-icons' }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$Cyan = [System.Drawing.Color]::FromArgb(255, 35, 186, 236)
$Ink = [System.Drawing.Color]::FromArgb(255, 175, 193, 205)
$RibbonInk = [System.Drawing.Color]::FromArgb(255, 48, 52, 59)
$script:K = 2
$Page = [System.Drawing.Color]::FromArgb(255, 17, 24, 32)
$Selected = [System.Drawing.Color]::FromArgb(255, 37, 67, 84)

# Ribbon artwork reused as-is (recolored).
$Reuse = @(
    'pptk-smart-bar', 'pptk-fix-colors', 'pptk-replace-colors', 'pptk-process', 'pptk-matrix', 'pptk-central',
    'pptk-tk-right', 'pptk-tk-bottom', 'pptk-tk-top', 'pptk-tk-left',
    'pptk-text', 'pptk-shape-rect', 'pptk-shape-round', 'pptk-shape-oval', 'pptk-shape-chevron',
    'pptk-shape-pentagon', 'pptk-shape-triangle', 'pptk-shape-para',
    'pptk-new-slide', 'pptk-insert-icon', 'pptk-enclose',
    'pptk-align-left', 'pptk-align-right', 'pptk-align-top', 'pptk-align-bottom', 'pptk-align-middle', 'pptk-align-center',
    'pptk-same-width', 'pptk-same-height', 'pptk-same-size', 'pptk-dist-h', 'pptk-dist-v',
    'pptk-swap', 'pptk-swap-text', 'pptk-merge-text', 'pptk-save-slides',
    'pptk-gap-h-inc', 'pptk-gap-h-dec', 'pptk-gap-v-inc', 'pptk-gap-v-dec',
    'pptk-copy-format', 'pptk-paste-format', 'pptk-paste-text', 'pptk-remove-margins'
)

function Brush($c) { New-Object System.Drawing.SolidBrush $c }
function Rect($g, $c, [single]$x, [single]$y, [single]$w, [single]$h) { $k = $script:K; $b = Brush $c; $g.FillRectangle($b, $x * $k, $y * $k, $w * $k, $h * $k); $b.Dispose() }
function Frame($g, $c, [single]$x, [single]$y, [single]$w, [single]$h) {
    # 1-unit outline inside the box.
    Rect $g $c $x $y $w 1; Rect $g $c $x ($y + $h - 1) $w 1
    Rect $g $c $x $y 1 $h; Rect $g $c ($x + $w - 1) $y 1 $h
}
function Poly($g, $c, [single[]]$pts) {
    $old = $g.SmoothingMode; $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $p = New-Object 'System.Drawing.PointF[]' ($pts.Count / 2)
    for ($i = 0; $i -lt $p.Count; $i++) { $p[$i] = New-Object System.Drawing.PointF ($pts[$i * 2] * $script:K), ($pts[$i * 2 + 1] * $script:K) }
    $b = Brush $c; $g.FillPolygon($b, $p); $b.Dispose(); $g.SmoothingMode = $old
}
function Ellipse($g, $c, [single]$x, [single]$y, [single]$w, [single]$h) {
    $old = $g.SmoothingMode; $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $k = $script:K; $b = Brush $c; $g.FillEllipse($b, $x * $k, $y * $k, $w * $k, $h * $k); $b.Dispose(); $g.SmoothingMode = $old
}
function Arc($g, $c, [single]$x, [single]$y, [single]$w, [single]$h, [single]$start, [single]$sweep, [single]$pen) {
    $old = $g.SmoothingMode; $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $k = $script:K; $p = New-Object System.Drawing.Pen $c, ($pen * $k); $g.DrawArc($p, $x * $k, $y * $k, $w * $k, $h * $k, $start, $sweep); $p.Dispose(); $g.SmoothingMode = $old
}
function Cross($g, $c, [single]$cx, [single]$cy, [single]$r) {
    $old = $g.SmoothingMode; $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $k = $script:K; $p = New-Object System.Drawing.Pen $c, (1.6 * $k)
    $g.DrawLine($p, ($cx - $r) * $k, ($cy - $r) * $k, ($cx + $r) * $k, ($cy + $r) * $k)
    $g.DrawLine($p, ($cx - $r) * $k, ($cy + $r) * $k, ($cx + $r) * $k, ($cy - $r) * $k)
    $p.Dispose(); $g.SmoothingMode = $old
}
function Clear-Cell($g, [single]$x, [single]$y, [single]$w, [single]$h) {
    $old = $g.CompositingMode; $g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
    Rect $g ([System.Drawing.Color]::Transparent) $x $y $w $h; $g.CompositingMode = $old
}

# Rows and columns: one family.
#   Keep bounds: two ink bracket lines mark the fixed outer edges.
#   Expand: an arrow points out of the grid. Shrink: an arrow points into it.
#   Add: the new column or row is solid cyan. Delete: it is gone and a cyan X marks it.
function Grid-Columns($g, [single]$x0, [bool]$added) {
    foreach ($r in 0, 1, 2) {
        Rect $g $Ink $x0 (2 + 4 * $r) 3 3
        Rect $g $Ink ($x0 + 4) (2 + 4 * $r) 3 3
        if ($added) { Rect $g $Cyan ($x0 + 8) (2 + 4 * $r) 3 3 }
    }
    if (-not $added) { Cross $g $Cyan ($x0 + 9.5) 7.5 2.5 }
}
function Grid-Rows($g, [single]$y0, [bool]$added) {
    foreach ($c in 0, 1, 2) {
        Rect $g $Ink (2 + 4 * $c) $y0 3 3
        Rect $g $Ink (2 + 4 * $c) ($y0 + 4) 3 3
        if ($added) { Rect $g $Cyan (2 + 4 * $c) ($y0 + 8) 3 3 }
    }
    if (-not $added) { Cross $g $Cyan 7.5 ($y0 + 9.5) 2.5 }
}
$Draw = [ordered]@{
    'sb-col-add-keep' = { param($g) Rect $g $Ink 0 1 1 14; Rect $g $Ink 15 1 1 14; Grid-Columns $g 2 $true }
    'sb-col-del-keep' = { param($g) Rect $g $Ink 0 1 1 14; Rect $g $Ink 15 1 1 14; Grid-Columns $g 2 $false }
    'sb-row-add-keep' = { param($g) Rect $g $Ink 1 0 14 1; Rect $g $Ink 1 15 14 1; Grid-Rows $g 2 $true }
    'sb-row-del-keep' = { param($g) Rect $g $Ink 1 0 14 1; Rect $g $Ink 1 15 14 1; Grid-Rows $g 2 $false }
    'sb-col-add-expand' = { param($g) Grid-Columns $g 0 $true; Poly $g $Cyan @(12, 5, 16, 8, 12, 11) }
    'sb-col-del-shrink' = { param($g) Grid-Columns $g 0 $false; Poly $g $Cyan @(16, 5, 12, 8, 16, 11) }
    'sb-row-add-expand' = { param($g) Grid-Rows $g 0 $true; Poly $g $Cyan @(5, 12, 8, 16, 11, 12) }
    'sb-row-del-shrink' = { param($g) Grid-Rows $g 0 $false; Poly $g $Cyan @(5, 16, 8, 12, 11, 16) }
    # Corner roundness: a square outline whose top-left corner is a thick cyan arc.
    'sb-corner' = { param($g)
        Rect $g $Ink 7 1 8 1; Rect $g $Ink 14 1 1 14; Rect $g $Ink 1 14 14 1; Rect $g $Ink 1 7 1 8
        Arc $g $Cyan 1.75 1.75 12 12 180 90 2 }
    # Unify shape type: a square and a triangle become two cyan circles.
    'sb-unify' = { param($g)
        Rect $g $Ink 1 1 5 5
        Poly $g $Ink @(3.5, 9, 6, 14, 1, 14)
        Poly $g $Cyan @(7, 2.5, 9, 3.5, 7, 4.5); Poly $g $Cyan @(7, 10.5, 9, 11.5, 7, 12.5)
        Ellipse $g $Cyan 10 1 5 5; Ellipse $g $Cyan 10 9 5 5 }
    # Flip: a slide outline with a dashed cyan mirror line and a shape on each side.
    'sb-flip-slide' = { param($g)
        Frame $g $Ink 0 3 16 10
        foreach ($y in 4, 7, 10) { Rect $g $Cyan 7.5 $y 1 2 }
        Poly $g $Cyan @(6, 5, 6, 11, 2, 8)
        Poly $g $Ink @(10, 5, 10, 11, 14, 8) }
    'sb-flip-deck' = { param($g)
        Rect $g $Ink 3 0 13 1; Rect $g $Ink 15 0 1 9
        Frame $g $Ink 0 3 14 10
        foreach ($y in 4, 7, 10) { Rect $g $Cyan 6.5 $y 1 2 }
        Poly $g $Cyan @(5, 5, 5, 11, 2, 8)
        Poly $g $Ink @(9, 5, 9, 11, 12, 8)
        Rect $g $Ink 3 14 13 1 }
    # Font size: a letter T with a cyan triangle up (larger) or down (smaller).
    'sb-font-up' = { param($g) Rect $g $Ink 0 3 10 2; Rect $g $Ink 4 3 2 12; Poly $g $Cyan @(13, 2, 16, 7, 10, 7) }
    'sb-font-down' = { param($g) Rect $g $Ink 1 6 8 2; Rect $g $Ink 4 6 2 9; Poly $g $Cyan @(10, 9, 16, 9, 13, 14) }
    # Stacking order: the cyan square is in front (bring forward) or behind (send backward).
    'sb-bring-forward' = { param($g)
        Rect $g $Ink 1 6 9 9
        Clear-Cell $g 4 3 10 10
        Rect $g $Cyan 5 4 9 9 }
    'sb-send-backward' = { param($g)
        Rect $g $Cyan 5 1 9 9
        Clear-Cell $g 1 5 10 10
        Rect $g $Ink 2 6 9 9 }
}

function New-Canvas([int]$size = 32) {
    $script:K = $size / 16
    $b = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($b)
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::None
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::Half
    return @($b, $g)
}

function Recolor-Ribbon([string]$name) {
    $src = New-Object System.Drawing.Bitmap (Join-Path $Ribbon32 "$name.png")
    $b = New-Object System.Drawing.Bitmap 32, 32, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    for ($y = 0; $y -lt 32; $y++) { for ($x = 0; $x -lt 32; $x++) {
        $p = $src.GetPixel($x, $y)
        if ($p.A -eq 0) { $b.SetPixel($x, $y, [System.Drawing.Color]::Transparent); continue }
        if ($p.R -gt 230 -and $p.G -gt 230 -and $p.B -gt 230) { $b.SetPixel($x, $y, [System.Drawing.Color]::Transparent); continue }
        $inkDistance = [Math]::Abs($p.R - 48) + [Math]::Abs($p.G - 52) + [Math]::Abs($p.B - 59)
        if ($inkDistance -lt 60) { $b.SetPixel($x, $y, [System.Drawing.Color]::FromArgb($p.A, $Ink.R, $Ink.G, $Ink.B)) } else { $b.SetPixel($x, $y, $p) }
    } }
    $src.Dispose()
    return $b
}

# Opaque 8-bit BMP, composited on one row color. A stored BMP keeps its exact
# colors in a form; a stored GIF loses transparency and a stored ICO is reduced
# to 16 colors. Each icon is written twice: on the normal row color and on the
# selected row color, and the Smart Bar shows the one that matches the row.
function Save-Bmp($bitmap, $back, [string]$path) {
    $colors = New-Object System.Collections.Generic.List[int]
    $index = New-Object 'byte[]' (32 * 32)
    for ($y = 0; $y -lt 32; $y++) { for ($x = 0; $x -lt 32; $x++) {
        $p = $bitmap.GetPixel($x, $y); $a = $p.A / 255.0
        $r = [int][Math]::Round($p.R * $a + $back.R * (1 - $a)); $g = [int][Math]::Round($p.G * $a + $back.G * (1 - $a)); $bl = [int][Math]::Round($p.B * $a + $back.B * (1 - $a))
        $key = ($r -shl 16) -bor ($g -shl 8) -bor $bl
        $at = $colors.IndexOf($key)
        if ($at -lt 0) { if ($colors.Count -ge 256) { throw "$path has more than 256 colors" }; $colors.Add($key); $at = $colors.Count - 1 }
        $index[$y * 32 + $x] = [byte]$at
    } }
    $out = New-Object System.Drawing.Bitmap 32, 32, ([System.Drawing.Imaging.PixelFormat]::Format8bppIndexed)
    $pal = $out.Palette
    for ($i = 0; $i -lt $pal.Entries.Length; $i++) {
        if ($i -lt $colors.Count) { $k = $colors[$i]; $pal.Entries[$i] = [System.Drawing.Color]::FromArgb(255, ($k -shr 16) -band 255, ($k -shr 8) -band 255, $k -band 255) }
        else { $pal.Entries[$i] = $back }
    }
    $out.Palette = $pal
    $rect = New-Object System.Drawing.Rectangle 0, 0, 32, 32
    $data = $out.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::WriteOnly, [System.Drawing.Imaging.PixelFormat]::Format8bppIndexed)
    for ($y = 0; $y -lt 32; $y++) { [System.Runtime.InteropServices.Marshal]::Copy($index, $y * 32, [IntPtr]($data.Scan0.ToInt64() + $y * $data.Stride), 32) }
    $out.UnlockBits($data)
    $out.Save($path, [System.Drawing.Imaging.ImageFormat]::Bmp)
    $out.Dispose()
}

if ($Ribbon) {
    $Ink = $RibbonInk
    foreach ($name in $Draw.Keys) {
        foreach ($size in 16, 32) {
            $c = New-Canvas $size; & $Draw[$name] $c[1]; $c[1].Dispose()
            $c[0].Save((Join-Path $Root "icons\$size\$name.png"), [System.Drawing.Imaging.ImageFormat]::Png); $c[0].Dispose()
        }
    }
    "wrote $($Draw.Count) ribbon icons to icons\16 and icons\32"
    exit 0
}

$built = [ordered]@{}
foreach ($name in $Reuse) { $built[$name] = Recolor-Ribbon $name }
foreach ($name in $Draw.Keys) { $c = New-Canvas; & $Draw[$name] $c[1]; $c[1].Dispose(); $built[$name] = $c[0] }
foreach ($name in $built.Keys) {
    Save-Bmp $built[$name] $Page (Join-Path $OutDir "$name.bmp")
    Save-Bmp $built[$name] $Selected (Join-Path $OutDir "$name-sel.bmp")
}

# Contact sheet: every icon on the normal and the selected row color, at 2x.
$names = @($built.Keys); $cols = 6; $cellW = 200; $cellH = 84
$rows = [Math]::Ceiling($names.Count / $cols)
$sheet = New-Object System.Drawing.Bitmap ($cols * $cellW), ($rows * $cellH), ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$sg = [System.Drawing.Graphics]::FromImage($sheet); $sg.Clear($Page)
$sg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
$font = New-Object System.Drawing.Font 'Segoe UI', 8; $text = Brush ([System.Drawing.Color]::FromArgb(255, 240, 245, 249))
$sel = Brush $Selected
for ($i = 0; $i -lt $names.Count; $i++) {
    $x = ($i % $cols) * $cellW; $y = [Math]::Floor($i / $cols) * $cellH
    $normal = [System.Drawing.Image]::FromFile((Join-Path $OutDir "$($names[$i]).bmp"))
    $chosen = [System.Drawing.Image]::FromFile((Join-Path $OutDir "$($names[$i])-sel.bmp"))
    $sg.DrawImage($normal, $x + 8, $y + 6, 32, 32)
    $sg.FillRectangle($sel, $x + 48, $y + 2, 40, 40); $sg.DrawImage($chosen, $x + 52, $y + 6, 32, 32)
    $normal.Dispose(); $chosen.Dispose()
    $sg.DrawString($names[$i], $font, $text, $x + 4, $y + 50)
}
$sg.Dispose(); $sheet.Save((Join-Path $OutDir 'contact-sheet.png'), [System.Drawing.Imaging.ImageFormat]::Png); $sheet.Dispose()
foreach ($b in $built.Values) { $b.Dispose() }
"built $($built.Count) icons in $OutDir"
