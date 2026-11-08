# Builds the PPTK cyan ribbon icon set and the integration manifest.
# Does not edit CustomUI.xml, the PPTM, forms, or command modules.
param(
    [string]$OnlyIcon = ''
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Icon16 = Join-Path $Root 'icons\16'
$Icon32 = Join-Path $Root 'icons\32'
New-Item -ItemType Directory -Force -Path $Icon16, $Icon32 | Out-Null

$script:Cyan = [System.Drawing.Color]::FromArgb(255, 35, 186, 236)
$script:Ink = [System.Drawing.Color]::FromArgb(255, 48, 52, 59)
$script:White = [System.Drawing.Color]::FromArgb(255, 255, 255)
$script:Light = [System.Drawing.Color]::FromArgb(255, 247, 249, 251)
$script:Dark = [System.Drawing.Color]::FromArgb(255, 48, 52, 59)

function Fill-Rect {
    param($Graphics, [single]$Scale, [single]$X, [single]$Y, [single]$W, [single]$H, [System.Drawing.Color]$Color)
    $brush = New-Object System.Drawing.SolidBrush $Color
    $Graphics.FillRectangle($brush, ($X * $Scale), ($Y * $Scale), ($W * $Scale), ($H * $Scale))
    $brush.Dispose()
}

function Stroke-Rect {
    param($Graphics, [single]$Scale, [single]$X, [single]$Y, [single]$W, [single]$H, [System.Drawing.Color]$Color, [single]$PenUnits = 1)
    $pen = New-Object System.Drawing.Pen $Color, ($PenUnits * $Scale)
    $inset = ($PenUnits * $Scale) / 2
    $Graphics.DrawRectangle($pen, ($X * $Scale) + $inset, ($Y * $Scale) + $inset, ($W * $Scale) - ($PenUnits * $Scale), ($H * $Scale) - ($PenUnits * $Scale))
    $pen.Dispose()
}

function Fill-Ellipse {
    param($Graphics, [single]$Scale, [single]$X, [single]$Y, [single]$W, [single]$H, [System.Drawing.Color]$Color)
    $old = $Graphics.SmoothingMode
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $brush = New-Object System.Drawing.SolidBrush $Color
    $Graphics.FillEllipse($brush, ($X * $Scale), ($Y * $Scale), ($W * $Scale), ($H * $Scale))
    $brush.Dispose()
    $Graphics.SmoothingMode = $old
}

function Stroke-Ellipse {
    param($Graphics, [single]$Scale, [single]$X, [single]$Y, [single]$W, [single]$H, [System.Drawing.Color]$Color, [single]$PenUnits = 1)
    $old = $Graphics.SmoothingMode
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $pen = New-Object System.Drawing.Pen $Color, ($PenUnits * $Scale)
    $inset = ($PenUnits * $Scale) / 2
    $Graphics.DrawEllipse($pen, ($X * $Scale) + $inset, ($Y * $Scale) + $inset, ($W * $Scale) - ($PenUnits * $Scale), ($H * $Scale) - ($PenUnits * $Scale))
    $pen.Dispose()
    $Graphics.SmoothingMode = $old
}

function Get-RoundPath {
    param([single]$X, [single]$Y, [single]$W, [single]$H, [single]$R)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = [Math]::Min($R * 2, [Math]::Min($W, $H))
    $path.AddArc($X, $Y, $d, $d, 180, 90)
    $path.AddArc(($X + $W - $d), $Y, $d, $d, 270, 90)
    $path.AddArc(($X + $W - $d), ($Y + $H - $d), $d, $d, 0, 90)
    $path.AddArc($X, ($Y + $H - $d), $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

function Fill-Round {
    param($Graphics, [single]$Scale, [single]$X, [single]$Y, [single]$W, [single]$H, [single]$R, [System.Drawing.Color]$Color)
    $old = $Graphics.SmoothingMode
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $path = Get-RoundPath ($X * $Scale) ($Y * $Scale) ($W * $Scale) ($H * $Scale) ($R * $Scale)
    $brush = New-Object System.Drawing.SolidBrush $Color
    $Graphics.FillPath($brush, $path)
    $brush.Dispose()
    $path.Dispose()
    $Graphics.SmoothingMode = $old
}

function Stroke-Round {
    param($Graphics, [single]$Scale, [single]$X, [single]$Y, [single]$W, [single]$H, [single]$R, [System.Drawing.Color]$Color, [single]$PenUnits = 1)
    $old = $Graphics.SmoothingMode
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $inset = ($PenUnits * $Scale) / 2
    $path = Get-RoundPath (($X * $Scale) + $inset) (($Y * $Scale) + $inset) (($W * $Scale) - (2 * $inset)) (($H * $Scale) - (2 * $inset)) ($R * $Scale)
    $pen = New-Object System.Drawing.Pen $Color, ($PenUnits * $Scale)
    $Graphics.DrawPath($pen, $path)
    $pen.Dispose()
    $path.Dispose()
    $Graphics.SmoothingMode = $old
}

function Fill-Poly {
    param($Graphics, [single]$Scale, [System.Drawing.Color]$Color, [single[]]$Coords)
    $count = [int]($Coords.Length / 2)
    $points = New-Object 'System.Drawing.PointF[]' $count
    for ($i = 0; $i -lt $count; $i++) {
        $points[$i] = New-Object System.Drawing.PointF ([single]($Coords[$i * 2] * $Scale)), ([single]($Coords[($i * 2) + 1] * $Scale))
    }
    $old = $Graphics.SmoothingMode
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $brush = New-Object System.Drawing.SolidBrush $Color
    $Graphics.FillPolygon($brush, $points)
    $brush.Dispose()
    $Graphics.SmoothingMode = $old
}

function Stroke-Line {
    param($Graphics, [single]$Scale, [single]$X1, [single]$Y1, [single]$X2, [single]$Y2, [System.Drawing.Color]$Color, [single]$PenUnits = 1.5)
    $old = $Graphics.SmoothingMode
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $pen = New-Object System.Drawing.Pen $Color, ($PenUnits * $Scale)
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $Graphics.DrawLine($pen, ($X1 * $Scale), ($Y1 * $Scale), ($X2 * $Scale), ($Y2 * $Scale))
    $pen.Dispose()
    $Graphics.SmoothingMode = $old
}

function Fill-Arrow {
    param($Graphics, [single]$Scale, [single]$X, [single]$Y, [single]$W, [single]$H, [string]$Direction, [System.Drawing.Color]$Color)
    $locals = @(
        @(0.00, 0.34), @(0.48, 0.34), @(0.48, 0.02),
        @(1.00, 0.50),
        @(0.48, 0.98), @(0.48, 0.66), @(0.00, 0.66)
    )
    $points = New-Object 'System.Drawing.PointF[]' $locals.Count
    for ($i = 0; $i -lt $locals.Count; $i++) {
        $px = [single]$locals[$i][0]
        $py = [single]$locals[$i][1]
        $dx = $px - 0.5
        $dy = $py - 0.5
        if ($Direction -eq 'R') { $nx = $px; $ny = $py }
        elseif ($Direction -eq 'L') { $nx = 1 - $px; $ny = $py }
        elseif ($Direction -eq 'D') { $nx = 0.5 - $dy; $ny = 0.5 + $dx }
        elseif ($Direction -eq 'U') { $nx = 0.5 + $dy; $ny = 0.5 - $dx }
        else { throw "Unknown arrow direction $Direction" }
        $points[$i] = New-Object System.Drawing.PointF ([single](($X + ($nx * $W)) * $Scale)), ([single](($Y + ($ny * $H)) * $Scale))
    }
    $old = $Graphics.SmoothingMode
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::None
    $brush = New-Object System.Drawing.SolidBrush $Color
    $Graphics.FillPolygon($brush, $points)
    $brush.Dispose()
    $Graphics.SmoothingMode = $old
}

function Draw-DoubleArrowH {
    param($Graphics, [single]$Scale, [single]$X, [single]$Y, [single]$W, [single]$H)
    $half = $W / 2
    Fill-Arrow $Graphics $Scale $X $Y $half $H 'L' $script:Cyan
    Fill-Arrow $Graphics $Scale ($X + $half) $Y $half $H 'R' $script:Cyan
}

function Draw-DoubleArrowV {
    param($Graphics, [single]$Scale, [single]$X, [single]$Y, [single]$W, [single]$H)
    $half = $H / 2
    Fill-Arrow $Graphics $Scale $X $Y $W $half 'U' $script:Cyan
    Fill-Arrow $Graphics $Scale $X ($Y + $half) $W $half 'D' $script:Cyan
}

function Icon-FormatMore($g, $k) {
    Fill-Rect $g $k 1 2 8 8 $script:Cyan
    Stroke-Rect $g $k 1 2 8 8 $script:Ink 1
    Fill-Ellipse $g $k 7 7 7 7 $script:Cyan
    Stroke-Ellipse $g $k 7 7 7 7 $script:Ink 1
}
function Icon-FixColors($g, $k) {
    Fill-Ellipse $g $k 1 6 3 4 $script:Ink
    Fill-Ellipse $g $k 6 4 4 8 $script:Cyan
    Fill-Ellipse $g $k 12 6 3 4 $script:Ink
}
function Icon-CopyFormat($g, $k) {
    Fill-Poly $g $k $script:Ink ([single[]]@(4, 8, 7, 11, 13, 4, 10, 1))
    Fill-Poly $g $k $script:Cyan ([single[]]@(1, 11, 6, 15, 8, 11, 3, 7))
}
function Icon-PasteFormat($g, $k) {
    Fill-Round $g $k 3 4 10 11 1.5 $script:Cyan
    Fill-Rect $g $k 6 1 4 4 $script:Ink
    Fill-Rect $g $k 5 8 6 1.4 $script:White
    Fill-Rect $g $k 5 11 6 1.4 $script:White
}
function Icon-SaveSlides($g, $k) {
    Fill-Rect $g $k 1 1 9 8 $script:Ink
    Fill-Rect $g $k 5 6 10 8 $script:Cyan
    Stroke-Rect $g $k 5 6 10 8 $script:Ink 1
    Fill-Rect $g $k 7 8 6 1.3 $script:White
    Fill-Rect $g $k 7 11 4 1.3 $script:White
}
function Icon-EmailSlides($g, $k) {
    Fill-Rect $g $k 3 1 7 5 $script:Ink
    Fill-Rect $g $k 4 2.1 5 0.9 $script:White
    Fill-Rect $g $k 4 3.5 3.4 0.9 $script:White
    Fill-Round $g $k 1 6 14 8 1.2 $script:Cyan
    Stroke-Line $g $k 2.2 7.2 8 11.2 $script:White 1.2
    Stroke-Line $g $k 13.8 7.2 8 11.2 $script:White 1.2
}
function Icon-LibrarySave($g, $k) {
    Fill-Rect $g $k 1 1 9 7 $script:Ink
    Fill-Rect $g $k 2.2 2.2 6.4 0.9 $script:White
    Fill-Rect $g $k 2.2 3.6 4 0.9 $script:White
    Fill-Round $g $k 5 8 10 7 1.2 $script:Cyan
    Fill-Rect $g $k 9.1 9.2 1.6 4.4 $script:White
    Fill-Rect $g $k 7.7 10.6 4.4 1.6 $script:White
}
function Icon-LibraryOpen($g, $k) {
    Fill-Rect $g $k 1 3 10 8 $script:Ink
    Fill-Rect $g $k 4 1 11 8 $script:Cyan
    Stroke-Rect $g $k 4 1 11 8 $script:Ink 1
    Fill-Rect $g $k 5.4 2.6 7 1 $script:White
    Fill-Rect $g $k 5.4 4.4 5 1 $script:White
}
function Icon-Text($g, $k) {
    Fill-Rect $g $k 2 2 12 3 $script:Cyan
    Fill-Rect $g $k 6 2 4 12 $script:Cyan
}
function Icon-NewSlide($g, $k) {
    Fill-Rect $g $k 1 1 9 9 $script:Ink
    Fill-Rect $g $k 3 3 5 1.3 $script:White
    Fill-Rect $g $k 3 5.4 4 1.3 $script:White
    Fill-Ellipse $g $k 8 8 7 7 $script:Cyan
    Fill-Rect $g $k 10.6 9.4 1.8 4.2 $script:White
    Fill-Rect $g $k 9.4 10.6 4.2 1.8 $script:White
}
function Icon-InsertIcon($g, $k) {
    Fill-Round $g $k 2 2 12 12 3 $script:Cyan
    Fill-Poly $g $k $script:White ([single[]]@(8, 4, 9.1, 6.9, 12, 8, 9.1, 9.1, 8, 12, 6.9, 9.1, 4, 8, 6.9, 6.9))
}
function Icon-Enclose($g, $k) {
    Stroke-Ellipse $g $k 2 2 12 12 $script:Cyan 2
    Fill-Poly $g $k $script:Ink ([single[]]@(8, 5, 11, 8, 8, 11, 5, 8))
}
function Icon-InsertMore($g, $k) {
    Fill-Rect $g $k 2 6 12 4 $script:Cyan
    Fill-Rect $g $k 6 2 4 12 $script:Cyan
}
function Draw-Tk($g, $k, $side) {
    if ($side -eq 'right') {
        Fill-Rect $g $k 1 4 7 8 $script:Ink
        Fill-Round $g $k 9 4 6 8 1.2 $script:Cyan
        Fill-Rect $g $k 10.2 5.4 3.6 1.3 $script:White
        Fill-Rect $g $k 11.3 5.4 1.4 4.6 $script:White
    } elseif ($side -eq 'left') {
        Fill-Round $g $k 1 4 6 8 1.2 $script:Cyan
        Fill-Rect $g $k 8 4 7 8 $script:Ink
        Fill-Rect $g $k 2.2 5.4 3.6 1.3 $script:White
        Fill-Rect $g $k 3.3 5.4 1.4 4.6 $script:White
    } elseif ($side -eq 'top') {
        Fill-Round $g $k 4 1 8 6 1.2 $script:Cyan
        Fill-Rect $g $k 4 8 8 7 $script:Ink
        Fill-Rect $g $k 5.4 2.2 5.2 1.3 $script:White
        Fill-Rect $g $k 7.3 2.2 1.4 3.6 $script:White
    } else {
        Fill-Rect $g $k 4 1 8 7 $script:Ink
        Fill-Round $g $k 4 9 8 6 1.2 $script:Cyan
        Fill-Rect $g $k 5.4 10.2 5.2 1.3 $script:White
        Fill-Rect $g $k 7.3 10.2 1.4 3.4 $script:White
    }
}
function Icon-TkRight($g, $k) { Draw-Tk $g $k 'right' }
function Icon-TkLeft($g, $k) { Draw-Tk $g $k 'left' }
function Icon-TkTop($g, $k) { Draw-Tk $g $k 'top' }
function Icon-TkBottom($g, $k) { Draw-Tk $g $k 'bottom' }
function Icon-TextMore($g, $k) {
    Fill-Rect $g $k 2 3 12 2 $script:Ink
    Fill-Rect $g $k 2 7 9 2 $script:Ink
    Fill-Rect $g $k 2 11 5 2 $script:Ink
    Fill-Rect $g $k 11 10 4 4 $script:Cyan
}
function Icon-SpecialChar($g, $k) {
    Stroke-Line $g $k 8 2 8 14 $script:Cyan 2
    Stroke-Line $g $k 2 8 14 8 $script:Cyan 2
    Stroke-Line $g $k 4 4 12 12 $script:Cyan 2
    Stroke-Line $g $k 12 4 4 12 $script:Cyan 2
}
function Icon-PasteText($g, $k) {
    Stroke-Round $g $k 3 4 10 11 1.5 $script:Ink 1.4
    Fill-Rect $g $k 6 1 4 4 $script:Cyan
    Fill-Rect $g $k 5 8 6 1.4 $script:Cyan
    Fill-Rect $g $k 5 11 5 1.4 $script:Cyan
}
function Icon-MergeText($g, $k) {
    Fill-Rect $g $k 1 1 6 4 $script:Ink
    Fill-Rect $g $k 9 1 6 4 $script:Ink
    Fill-Arrow $g $k 6 5.2 4 4 'D' $script:Cyan
    Fill-Rect $g $k 2 10 12 4 $script:Cyan
}
function Icon-RemoveMargins($g, $k) {
    Stroke-Rect $g $k 2 2 12 12 $script:Ink 1.3
    Fill-Rect $g $k 3 5 10 1.5 $script:Cyan
    Fill-Rect $g $k 3 8 10 1.5 $script:Cyan
    Fill-Rect $g $k 3 11 10 1.5 $script:Cyan
}
function Icon-AlignLeft($g, $k) {
    Fill-Rect $g $k 1 1 3 14 $script:Cyan
    Fill-Rect $g $k 6 3 4 2 $script:Ink
    Fill-Rect $g $k 6 7 8 2 $script:Ink
    Fill-Rect $g $k 6 11 6 2 $script:Ink
}
function Icon-AlignCenter($g, $k) {
    Fill-Rect $g $k 2 3 12 2 $script:Ink
    Fill-Rect $g $k 4 7 8 2 $script:Ink
    Fill-Rect $g $k 3 11 10 2 $script:Ink
    Fill-Rect $g $k 7 1 2 14 $script:Cyan
}
function Icon-AlignRight($g, $k) {
    Fill-Rect $g $k 12 1 3 14 $script:Cyan
    Fill-Rect $g $k 6 3 4 2 $script:Ink
    Fill-Rect $g $k 2 7 8 2 $script:Ink
    Fill-Rect $g $k 4 11 6 2 $script:Ink
}
function Icon-AlignTop($g, $k) {
    Fill-Rect $g $k 1 1 14 3 $script:Cyan
    Fill-Rect $g $k 3 6 2 4 $script:Ink
    Fill-Rect $g $k 7 6 2 8 $script:Ink
    Fill-Rect $g $k 11 6 2 6 $script:Ink
}
function Icon-AlignMiddle($g, $k) {
    Fill-Rect $g $k 3 3 2 10 $script:Ink
    Fill-Rect $g $k 7 1 2 14 $script:Ink
    Fill-Rect $g $k 11 4 2 8 $script:Ink
    Fill-Rect $g $k 1 7 14 2 $script:Cyan
}
function Icon-AlignBottom($g, $k) {
    Fill-Rect $g $k 1 12 14 3 $script:Cyan
    Fill-Rect $g $k 3 6 2 4 $script:Ink
    Fill-Rect $g $k 7 2 2 8 $script:Ink
    Fill-Rect $g $k 11 4 2 6 $script:Ink
}
function Icon-DistH($g, $k) {
    Fill-Rect $g $k 1 5 3 6 $script:Ink
    Fill-Rect $g $k 7 5 3 6 $script:Cyan
    Fill-Rect $g $k 13 5 3 6 $script:Ink
}
function Icon-DistV($g, $k) {
    Fill-Rect $g $k 5 1 6 3 $script:Ink
    Fill-Rect $g $k 5 7 6 3 $script:Cyan
    Fill-Rect $g $k 5 13 6 3 $script:Ink
}
function Icon-ResizeSpace($g, $k) {
    Fill-Rect $g $k 1 2 3 3 $script:Ink
    Fill-Rect $g $k 6.5 2 3 3 $script:Ink
    Fill-Rect $g $k 12 2 3 3 $script:Ink
    Draw-DoubleArrowH $g $k 1 8 14 6
}
function Icon-ResizeH($g, $k) {
    Fill-Rect $g $k 1 2 3 8 $script:Ink
    Fill-Rect $g $k 6.5 2 3 8 $script:Ink
    Fill-Rect $g $k 12 2 3 8 $script:Ink
    Draw-DoubleArrowH $g $k 1 11 14 4
}
function Icon-ResizeHFirst($g, $k) {
    Fill-Rect $g $k 1 2 3 8 $script:Cyan
    Fill-Rect $g $k 6.5 2 3 8 $script:Ink
    Fill-Rect $g $k 12 2 3 8 $script:Ink
    Draw-DoubleArrowH $g $k 6 11 9 4
}
function Icon-ResizeHLast($g, $k) {
    Fill-Rect $g $k 1 2 3 8 $script:Ink
    Fill-Rect $g $k 6.5 2 3 8 $script:Ink
    Fill-Rect $g $k 12 2 3 8 $script:Cyan
    Draw-DoubleArrowH $g $k 1 11 9 4
}
function Icon-ResizeV($g, $k) {
    Fill-Rect $g $k 1 1 9 3 $script:Ink
    Fill-Rect $g $k 1 6.5 9 3 $script:Ink
    Fill-Rect $g $k 1 12 9 3 $script:Ink
    Draw-DoubleArrowV $g $k 11 1 4 14
}
function Icon-ResizeVFirst($g, $k) {
    Fill-Rect $g $k 2 1 8 3 $script:Cyan
    Fill-Rect $g $k 2 6.5 8 3 $script:Ink
    Fill-Rect $g $k 2 12 8 3 $script:Ink
    Draw-DoubleArrowV $g $k 11 6 4 9
}
function Icon-ResizeVLast($g, $k) {
    Fill-Rect $g $k 2 1 8 3 $script:Ink
    Fill-Rect $g $k 2 6.5 8 3 $script:Ink
    Fill-Rect $g $k 2 12 8 3 $script:Cyan
    Draw-DoubleArrowV $g $k 11 1 4 9
}
function Icon-Gaps($g, $k) {
    Fill-Rect $g $k 1 3 5 10 $script:Ink
    Fill-Rect $g $k 10 3 5 10 $script:Ink
    Fill-Rect $g $k 7 3 2 10 $script:Cyan
}
function Icon-GapHLeft($g, $k) {
    Fill-Rect $g $k 1 4 4 8 $script:Ink
    Fill-Rect $g $k 11 4 4 8 $script:Ink
    Fill-Arrow $g $k 5 4 6 8 'L' $script:Cyan
}
function Icon-GapHRight($g, $k) {
    Fill-Rect $g $k 1 4 4 8 $script:Ink
    Fill-Rect $g $k 11 4 4 8 $script:Ink
    Fill-Arrow $g $k 5 4 6 8 'R' $script:Cyan
}
function Icon-GapHInc($g, $k) {
    Fill-Rect $g $k 1 4 3 8 $script:Ink
    Fill-Rect $g $k 12 4 3 8 $script:Ink
    Fill-Arrow $g $k 4 4 4 8 'L' $script:Cyan
    Fill-Arrow $g $k 8 4 4 8 'R' $script:Cyan
}
function Icon-GapHDec($g, $k) {
    Fill-Rect $g $k 1 4 3 8 $script:Ink
    Fill-Rect $g $k 12 4 3 8 $script:Ink
    Fill-Arrow $g $k 4 4 4 8 'R' $script:Cyan
    Fill-Arrow $g $k 8 4 4 8 'L' $script:Cyan
}
function Icon-GapVUp($g, $k) {
    Fill-Rect $g $k 4 1 8 4 $script:Ink
    Fill-Rect $g $k 4 11 8 4 $script:Ink
    Fill-Arrow $g $k 4 5 8 6 'U' $script:Cyan
}
function Icon-GapVDown($g, $k) {
    Fill-Rect $g $k 4 1 8 4 $script:Ink
    Fill-Rect $g $k 4 11 8 4 $script:Ink
    Fill-Arrow $g $k 4 5 8 6 'D' $script:Cyan
}
function Icon-GapVInc($g, $k) {
    Fill-Rect $g $k 4 1 8 3 $script:Ink
    Fill-Rect $g $k 4 12 8 3 $script:Ink
    Fill-Arrow $g $k 4 4 8 4 'U' $script:Cyan
    Fill-Arrow $g $k 4 8 8 4 'D' $script:Cyan
}
function Icon-GapVDec($g, $k) {
    Fill-Rect $g $k 4 1 8 3 $script:Ink
    Fill-Rect $g $k 4 12 8 3 $script:Ink
    Fill-Arrow $g $k 4 4 8 4 'D' $script:Cyan
    Fill-Arrow $g $k 4 8 8 4 'U' $script:Cyan
}
function Icon-ChangeShape($g, $k) {
    Fill-Rect $g $k 1 2 7 7 $script:Cyan
    Stroke-Rect $g $k 1 2 7 7 $script:Ink 1
    Fill-Ellipse $g $k 8 7 7 7 $script:Cyan
    Stroke-Ellipse $g $k 8 7 7 7 $script:Ink 1
}
function Icon-SameWidth($g, $k) {
    Fill-Rect $g $k 2 1 4 8 $script:Ink
    Fill-Rect $g $k 10 3 4 6 $script:Ink
    Draw-DoubleArrowH $g $k 1 11 14 4
}
function Icon-SameHeight($g, $k) {
    Fill-Rect $g $k 1 2 9 3 $script:Ink
    Fill-Rect $g $k 1 8 9 3 $script:Ink
    Draw-DoubleArrowV $g $k 11 1 4 14
}
function Icon-SameSize($g, $k) {
    Stroke-Rect $g $k 3 3 10 10 $script:Cyan 2
    Fill-Rect $g $k 7 7 2 2 $script:Ink
}
function Icon-Swap($g, $k) {
    Fill-Arrow $g $k 1 1 14 6 'R' $script:Cyan
    Fill-Arrow $g $k 1 9 14 6 'L' $script:Cyan
}
function Icon-SwapCenter($g, $k) {
    Fill-Arrow $g $k 1 1 14 5 'R' $script:Cyan
    Fill-Arrow $g $k 1 10 14 5 'L' $script:Cyan
    Fill-Rect $g $k 6 7 4 2 $script:Ink
    Fill-Rect $g $k 7 6 2 4 $script:Ink
}
function Icon-SwapText($g, $k) {
    Fill-Rect $g $k 1 1 5 2 $script:Ink
    Fill-Rect $g $k 2.5 1 2 6 $script:Ink
    Fill-Rect $g $k 10 1 5 2 $script:Cyan
    Fill-Rect $g $k 11.5 1 2 6 $script:Cyan
    Draw-DoubleArrowH $g $k 1 10 14 5
}
function Icon-ShapeMore($g, $k) {
    Fill-Rect $g $k 1 6 4 4 $script:Cyan
    Fill-Ellipse $g $k 6 6 4 4 $script:Ink
    Fill-Poly $g $k $script:Cyan ([single[]]@(12, 10, 13.5, 6, 15, 10))
}
function Icon-SizeTall($g, $k) {
    Fill-Rect $g $k 2 8 4 6 $script:Ink
    Fill-Rect $g $k 9 2 4 12 $script:Cyan
}
function Icon-SizeShort($g, $k) {
    Fill-Rect $g $k 2 2 4 12 $script:Ink
    Fill-Rect $g $k 9 8 4 6 $script:Cyan
}
function Icon-SizeWide($g, $k) {
    Fill-Rect $g $k 2 2 12 4 $script:Cyan
    Fill-Rect $g $k 4 10 6 4 $script:Ink
}
function Icon-SizeNarrow($g, $k) {
    Fill-Rect $g $k 2 2 12 4 $script:Ink
    Fill-Rect $g $k 5 10 6 4 $script:Cyan
}
function Icon-ColAdd($g, $k) {
    Fill-Rect $g $k 1 2 4 4 $script:Ink
    Fill-Rect $g $k 6 2 4 4 $script:Ink
    Fill-Rect $g $k 1 7 4 4 $script:Ink
    Fill-Rect $g $k 6 7 4 4 $script:Ink
    Fill-Rect $g $k 11 2 4 9 $script:Cyan
    Fill-Rect $g $k 12 4 2 5 $script:White
    Fill-Rect $g $k 11 6 4 2 $script:White
}
function Icon-RowAdd($g, $k) {
    Fill-Rect $g $k 2 1 4 4 $script:Ink
    Fill-Rect $g $k 7 1 4 4 $script:Ink
    Fill-Rect $g $k 2 6 4 4 $script:Ink
    Fill-Rect $g $k 7 6 4 4 $script:Ink
    Fill-Rect $g $k 2 11 9 4 $script:Cyan
    Fill-Rect $g $k 5.2 12 2.6 2 $script:White
    Fill-Rect $g $k 6 11.4 1.2 3.2 $script:White
}
function Icon-ColDel($g, $k) {
    Fill-Rect $g $k 1 2 4 4 $script:Ink
    Fill-Rect $g $k 6 2 4 4 $script:Ink
    Fill-Rect $g $k 1 7 4 4 $script:Ink
    Fill-Rect $g $k 6 7 4 4 $script:Ink
    Stroke-Rect $g $k 11 2 4 9 $script:Ink 1.2
    Fill-Rect $g $k 11 6 4 3 $script:Cyan
}
function Icon-RowDel($g, $k) {
    Fill-Rect $g $k 2 1 4 4 $script:Ink
    Fill-Rect $g $k 7 1 4 4 $script:Ink
    Fill-Rect $g $k 2 6 4 4 $script:Ink
    Fill-Rect $g $k 7 6 4 4 $script:Ink
    Stroke-Rect $g $k 2 11 9 4 $script:Ink 1.2
    Fill-Rect $g $k 4 11.5 5 3 $script:Cyan
}
function Icon-Arrange($g, $k) {
    Fill-Rect $g $k 2 2 5 5 $script:Cyan
    Fill-Rect $g $k 9 2 5 5 $script:Ink
    Fill-Rect $g $k 2 9 5 5 $script:Ink
    Fill-Rect $g $k 9 9 5 5 $script:Cyan
}
function Icon-Stretch($g, $k) {
    Fill-Rect $g $k 5 5 6 6 $script:Ink
    Fill-Arrow $g $k 0.4 6 4 4 'L' $script:Cyan
    Fill-Arrow $g $k 11.6 6 4 4 'R' $script:Cyan
    Fill-Arrow $g $k 6 0.4 4 4 'U' $script:Cyan
    Fill-Arrow $g $k 6 11.6 4 4 'D' $script:Cyan
}
function Icon-StretchLeft($g, $k) {
    Fill-Rect $g $k 9 3 6 10 $script:Ink
    Fill-Arrow $g $k 1 4 7 8 'L' $script:Cyan
}
function Icon-StretchRight($g, $k) {
    Fill-Rect $g $k 1 3 6 10 $script:Ink
    Fill-Arrow $g $k 8 4 7 8 'R' $script:Cyan
}
function Icon-StretchTop($g, $k) {
    Fill-Rect $g $k 3 9 10 6 $script:Ink
    Fill-Arrow $g $k 4 1 8 7 'U' $script:Cyan
}
function Icon-StretchBottom($g, $k) {
    Fill-Rect $g $k 3 1 10 6 $script:Ink
    Fill-Arrow $g $k 4 8 8 7 'D' $script:Cyan
}
function Icon-StretchToLeft($g, $k) {
    Fill-Rect $g $k 1 3 4 10 $script:Ink
    Fill-Rect $g $k 5 1 2 14 $script:Cyan
    Fill-Poly $g $k $script:Cyan ([single[]]@(7, 8, 11, 5, 11, 11))
    Fill-Rect $g $k 12 5 3 6 $script:Ink
}
function Icon-StretchToRight($g, $k) {
    Fill-Rect $g $k 11 3 4 10 $script:Ink
    Fill-Rect $g $k 9 1 2 14 $script:Cyan
    Fill-Poly $g $k $script:Cyan ([single[]]@(8, 8, 4, 5, 4, 11))
    Fill-Rect $g $k 1 5 3 6 $script:Ink
}
function Icon-StretchToTop($g, $k) {
    Fill-Rect $g $k 3 1 10 3 $script:Ink
    Fill-Rect $g $k 1 4 14 2 $script:Cyan
    Fill-Poly $g $k $script:Cyan ([single[]]@(5, 11, 11, 11, 8, 7))
    Fill-Rect $g $k 4 12 8 3 $script:Ink
}
function Icon-StretchToBottom($g, $k) {
    Fill-Rect $g $k 4 1 8 3 $script:Ink
    Fill-Poly $g $k $script:Cyan ([single[]]@(5, 5, 11, 5, 8, 9))
    Fill-Rect $g $k 1 10 14 2 $script:Cyan
    Fill-Rect $g $k 3 12 10 3 $script:Ink
}
function Icon-AlignMore($g, $k) {
    Fill-Rect $g $k 7 1 2 14 $script:Cyan
    Fill-Rect $g $k 1 7 14 2 $script:Cyan
    Fill-Rect $g $k 2 2 3 3 $script:Ink
    Fill-Rect $g $k 11 2 3 3 $script:Ink
    Fill-Rect $g $k 2 11 3 3 $script:Ink
    Fill-Rect $g $k 11 11 3 3 $script:Ink
}
function Icon-AlignAngle($g, $k) {
    Fill-Rect $g $k 1 7 5 5 $script:Ink
    Fill-Poly $g $k $script:Cyan ([single[]]@(9, 3, 14, 8, 9, 13, 5.5, 8))
}
function Icon-Process($g, $k) {
    Fill-Poly $g $k $script:Cyan ([single[]]@(1, 4, 3.2, 4, 5, 8, 3.2, 12, 1, 12, 2.8, 8))
    Fill-Poly $g $k $script:Cyan ([single[]]@(6, 4, 8.2, 4, 10, 8, 8.2, 12, 6, 12, 7.8, 8))
    Fill-Poly $g $k $script:Cyan ([single[]]@(11, 4, 13.2, 4, 15, 8, 13.2, 12, 11, 12, 12.8, 8))
}
function Icon-Matrix($g, $k) {
    Stroke-Rect $g $k 1 2 14 12 $script:Cyan 1.4
    foreach ($x in @(2.4, 6.6, 10.8)) {
        foreach ($y in @(3.2, 8.2)) {
            Fill-Rect $g $k $x $y 3 3.6 $script:Ink
        }
    }
}
function Icon-Central($g, $k) {
    Fill-Ellipse $g $k 5 5 6 6 $script:Cyan
    Fill-Rect $g $k 1 1 3 3 $script:Ink
    Fill-Rect $g $k 12 1 3 3 $script:Ink
    Fill-Rect $g $k 1 12 3 3 $script:Ink
    Fill-Rect $g $k 12 12 3 3 $script:Ink
}
function Icon-ShapeRect($g, $k) {
    Fill-Rect $g $k 2 4 12 8 $script:Cyan
    Stroke-Rect $g $k 2 4 12 8 $script:Ink 1
}
function Icon-ShapeRound($g, $k) {
    Fill-Round $g $k 2 4 12 8 2.5 $script:Cyan
    Stroke-Round $g $k 2 4 12 8 2.5 $script:Ink 1
}
function Icon-ShapeOval($g, $k) {
    Fill-Ellipse $g $k 2 4 12 8 $script:Cyan
    Stroke-Ellipse $g $k 2 4 12 8 $script:Ink 1
}
function Icon-ShapeChevron($g, $k) {
    Fill-Poly $g $k $script:Cyan ([single[]]@(1, 3, 9, 3, 14, 8, 9, 13, 1, 13, 6, 8))
}
function Icon-ShapePentagon($g, $k) {
    Fill-Poly $g $k $script:Cyan ([single[]]@(8, 1, 15, 6, 12.5, 14, 3.5, 14, 1, 6))
}
function Icon-ShapeTriangle($g, $k) {
    Fill-Poly $g $k $script:Cyan ([single[]]@(8, 1, 15, 14, 1, 14))
}
function Icon-ShapePara($g, $k) {
    Fill-Poly $g $k $script:Cyan ([single[]]@(4, 3, 15, 3, 12, 13, 1, 13))
}
function Icon-SmartBar($g, $k) {
    Fill-Round $g $k 1 4 14 8 3 $script:Cyan
    Fill-Ellipse $g $k 8.2 6 3.6 3.6 $script:White
    Stroke-Line $g $k 11.2 9.2 13.4 11.6 $script:Ink 1.6
}
function Icon-Table($g, $k) {
    Fill-Rect $g $k 2 2 12 4 $script:Cyan
    Stroke-Rect $g $k 2 2 12 12 $script:Ink 1.2
    Fill-Rect $g $k 2 8 12 1.2 $script:Ink
    Fill-Rect $g $k 8 2 1.2 12 $script:Ink
}
function Icon-Tools($g, $k) {
    Fill-Rect $g $k 2 3 12 2 $script:Ink
    Fill-Rect $g $k 2 7 12 2 $script:Ink
    Fill-Rect $g $k 2 11 12 2 $script:Ink
    Fill-Rect $g $k 2 2 4 4 $script:Cyan
    Fill-Rect $g $k 7 6 4 4 $script:Cyan
    Fill-Rect $g $k 10 10 4 4 $script:Cyan
}
function Icon-Storyline($g, $k) {
    Fill-Rect $g $k 3 1 10 7 $script:Ink
    Fill-Rect $g $k 5 3 6 1.2 $script:White
    Fill-Rect $g $k 2 9 12 5 $script:Cyan
    Fill-Rect $g $k 4 10.6 8 1.3 $script:White
}

$Draw = @{
    'pptk-format-more' = { param($g, $k) Icon-FormatMore $g $k }
    'pptk-fix-colors' = { param($g, $k) Icon-FixColors $g $k }
    'pptk-copy-format' = { param($g, $k) Icon-CopyFormat $g $k }
    'pptk-paste-format' = { param($g, $k) Icon-PasteFormat $g $k }
    'pptk-save-slides' = { param($g, $k) Icon-SaveSlides $g $k }
    'pptk-email-slides' = { param($g, $k) Icon-EmailSlides $g $k }
    'pptk-library-save' = { param($g, $k) Icon-LibrarySave $g $k }
    'pptk-library-open' = { param($g, $k) Icon-LibraryOpen $g $k }
    'pptk-text' = { param($g, $k) Icon-Text $g $k }
    'pptk-new-slide' = { param($g, $k) Icon-NewSlide $g $k }
    'pptk-insert-icon' = { param($g, $k) Icon-InsertIcon $g $k }
    'pptk-enclose' = { param($g, $k) Icon-Enclose $g $k }
    'pptk-insert-more' = { param($g, $k) Icon-InsertMore $g $k }
    'pptk-tk-right' = { param($g, $k) Icon-TkRight $g $k }
    'pptk-tk-left' = { param($g, $k) Icon-TkLeft $g $k }
    'pptk-tk-top' = { param($g, $k) Icon-TkTop $g $k }
    'pptk-tk-bottom' = { param($g, $k) Icon-TkBottom $g $k }
    'pptk-text-more' = { param($g, $k) Icon-TextMore $g $k }
    'pptk-special-char' = { param($g, $k) Icon-SpecialChar $g $k }
    'pptk-paste-text' = { param($g, $k) Icon-PasteText $g $k }
    'pptk-merge-text' = { param($g, $k) Icon-MergeText $g $k }
    'pptk-remove-margins' = { param($g, $k) Icon-RemoveMargins $g $k }
    'pptk-align-left' = { param($g, $k) Icon-AlignLeft $g $k }
    'pptk-align-center' = { param($g, $k) Icon-AlignCenter $g $k }
    'pptk-align-right' = { param($g, $k) Icon-AlignRight $g $k }
    'pptk-align-top' = { param($g, $k) Icon-AlignTop $g $k }
    'pptk-align-middle' = { param($g, $k) Icon-AlignMiddle $g $k }
    'pptk-align-bottom' = { param($g, $k) Icon-AlignBottom $g $k }
    'pptk-dist-h' = { param($g, $k) Icon-DistH $g $k }
    'pptk-dist-v' = { param($g, $k) Icon-DistV $g $k }
    'pptk-resize-space' = { param($g, $k) Icon-ResizeSpace $g $k }
    'pptk-resize-h' = { param($g, $k) Icon-ResizeH $g $k }
    'pptk-resize-h-first' = { param($g, $k) Icon-ResizeHFirst $g $k }
    'pptk-resize-h-last' = { param($g, $k) Icon-ResizeHLast $g $k }
    'pptk-resize-v' = { param($g, $k) Icon-ResizeV $g $k }
    'pptk-resize-v-first' = { param($g, $k) Icon-ResizeVFirst $g $k }
    'pptk-resize-v-last' = { param($g, $k) Icon-ResizeVLast $g $k }
    'pptk-gaps' = { param($g, $k) Icon-Gaps $g $k }
    'pptk-gap-h-left' = { param($g, $k) Icon-GapHLeft $g $k }
    'pptk-gap-h-right' = { param($g, $k) Icon-GapHRight $g $k }
    'pptk-gap-h-inc' = { param($g, $k) Icon-GapHInc $g $k }
    'pptk-gap-h-dec' = { param($g, $k) Icon-GapHDec $g $k }
    'pptk-gap-v-up' = { param($g, $k) Icon-GapVUp $g $k }
    'pptk-gap-v-down' = { param($g, $k) Icon-GapVDown $g $k }
    'pptk-gap-v-inc' = { param($g, $k) Icon-GapVInc $g $k }
    'pptk-gap-v-dec' = { param($g, $k) Icon-GapVDec $g $k }
    'pptk-change-shape' = { param($g, $k) Icon-ChangeShape $g $k }
    'pptk-same-width' = { param($g, $k) Icon-SameWidth $g $k }
    'pptk-same-height' = { param($g, $k) Icon-SameHeight $g $k }
    'pptk-same-size' = { param($g, $k) Icon-SameSize $g $k }
    'pptk-swap' = { param($g, $k) Icon-Swap $g $k }
    'pptk-swap-center' = { param($g, $k) Icon-SwapCenter $g $k }
    'pptk-swap-text' = { param($g, $k) Icon-SwapText $g $k }
    'pptk-shape-more' = { param($g, $k) Icon-ShapeMore $g $k }
    'pptk-size-tall' = { param($g, $k) Icon-SizeTall $g $k }
    'pptk-size-short' = { param($g, $k) Icon-SizeShort $g $k }
    'pptk-size-wide' = { param($g, $k) Icon-SizeWide $g $k }
    'pptk-size-narrow' = { param($g, $k) Icon-SizeNarrow $g $k }
    'pptk-col-add' = { param($g, $k) Icon-ColAdd $g $k }
    'pptk-row-add' = { param($g, $k) Icon-RowAdd $g $k }
    'pptk-col-del' = { param($g, $k) Icon-ColDel $g $k }
    'pptk-row-del' = { param($g, $k) Icon-RowDel $g $k }
    'pptk-arrange' = { param($g, $k) Icon-Arrange $g $k }
    'pptk-stretch' = { param($g, $k) Icon-Stretch $g $k }
    'pptk-stretch-left' = { param($g, $k) Icon-StretchLeft $g $k }
    'pptk-stretch-right' = { param($g, $k) Icon-StretchRight $g $k }
    'pptk-stretch-top' = { param($g, $k) Icon-StretchTop $g $k }
    'pptk-stretch-bottom' = { param($g, $k) Icon-StretchBottom $g $k }
    'pptk-stretch-to-left' = { param($g, $k) Icon-StretchToLeft $g $k }
    'pptk-stretch-to-right' = { param($g, $k) Icon-StretchToRight $g $k }
    'pptk-stretch-to-top' = { param($g, $k) Icon-StretchToTop $g $k }
    'pptk-stretch-to-bottom' = { param($g, $k) Icon-StretchToBottom $g $k }
    'pptk-align-more' = { param($g, $k) Icon-AlignMore $g $k }
    'pptk-align-angle' = { param($g, $k) Icon-AlignAngle $g $k }
    'pptk-process' = { param($g, $k) Icon-Process $g $k }
    'pptk-matrix' = { param($g, $k) Icon-Matrix $g $k }
    'pptk-central' = { param($g, $k) Icon-Central $g $k }
    'pptk-shape-rect' = { param($g, $k) Icon-ShapeRect $g $k }
    'pptk-shape-round' = { param($g, $k) Icon-ShapeRound $g $k }
    'pptk-shape-oval' = { param($g, $k) Icon-ShapeOval $g $k }
    'pptk-shape-chevron' = { param($g, $k) Icon-ShapeChevron $g $k }
    'pptk-shape-pentagon' = { param($g, $k) Icon-ShapePentagon $g $k }
    'pptk-shape-triangle' = { param($g, $k) Icon-ShapeTriangle $g $k }
    'pptk-shape-para' = { param($g, $k) Icon-ShapePara $g $k }
    'pptk-smart-bar' = { param($g, $k) Icon-SmartBar $g $k }
    'pptk-table' = { param($g, $k) Icon-Table $g $k }
    'pptk-tools' = { param($g, $k) Icon-Tools $g $k }
    'pptk-storyline' = { param($g, $k) Icon-Storyline $g $k }
}

# imageMso is the value on that control today. Empty means the control has no image.
# Action replace: remove imageMso and set image. Action add: set image only.
$Map = @(
    @{ Id = 'DesignMoreFormatting'; Group = 'Design'; Label = 'More Formatting'; Kind = 'menu'; Action = 'replace'; ImageMso = 'ShapeFillColorPicker'; Image = 'pptk-format-more' }
    @{ Id = 'PPTKFixSlideColors'; Group = 'Design'; Label = 'Fix Slide Colors'; Kind = 'button'; Action = 'replace'; ImageMso = 'ShapeFillColorPicker'; Image = 'pptk-fix-colors' }
    @{ Id = 'PPTKCopyFormatting'; Group = 'Design'; Label = 'Copy formatting'; Kind = 'button'; Action = 'replace'; ImageMso = 'FormatPainter'; Image = 'pptk-copy-format' }
    @{ Id = 'PPTKPasteFormatting'; Group = 'Design'; Label = 'Paste formatting'; Kind = 'button'; Action = 'replace'; ImageMso = 'FormatPainter'; Image = 'pptk-paste-format' }
    @{ Id = 'SaveSelectedSlidesButton'; Group = 'Reuse'; Label = 'Save selected slides'; Kind = 'button'; Action = 'replace'; ImageMso = 'FileSaveAs'; Image = 'pptk-save-slides' }
    @{ Id = 'EmailSelectedSlidesButton'; Group = 'Reuse'; Label = 'Email Slides'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-email-slides' }
    @{ Id = 'SaveToLibraryButton'; Group = 'Reuse'; Label = 'Save to Library'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-library-save' }
    @{ Id = 'OpenSlideLibraryButton'; Group = 'Reuse'; Label = 'Open Slide Library'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-library-open' }
    @{ Id = 'PPTKInsertT'; Group = 'Insert'; Label = 'Text'; Kind = 'button'; Action = 'replace'; ImageMso = 'TextBoxInsert'; Image = 'pptk-text' }
    @{ Id = 'PPTKInsertN'; Group = 'Insert'; Label = 'New slide'; Kind = 'button'; Action = 'replace'; ImageMso = 'SlideNew'; Image = 'pptk-new-slide' }
    @{ Id = 'PPTKInsertI'; Group = 'Insert'; Label = 'Icon'; Kind = 'button'; Action = 'replace'; ImageMso = 'IconInsertFromFile'; Image = 'pptk-insert-icon' }
    @{ Id = 'PPTKEncloseIcons'; Group = 'Insert'; Label = 'Enclose'; Kind = 'button'; Action = 'replace'; ImageMso = 'ShapeOval'; Image = 'pptk-enclose' }
    @{ Id = 'InsertMoreMenu'; Group = 'Insert'; Label = 'More'; Kind = 'menu'; Action = 'replace'; ImageMso = 'PasteAlternative'; Image = 'pptk-insert-more' }
    @{ Id = 'PPTKAddRightTK'; Group = 'Insert'; Label = 'Add Right TK'; Kind = 'button'; Action = 'replace'; ImageMso = 'TextBoxInsert'; Image = 'pptk-tk-right' }
    @{ Id = 'PPTKAddLeftTK'; Group = 'Insert'; Label = 'Add Left TK'; Kind = 'button'; Action = 'replace'; ImageMso = 'TextBoxInsert'; Image = 'pptk-tk-left' }
    @{ Id = 'PPTKAddTopTK'; Group = 'Insert'; Label = 'Add Top TK'; Kind = 'button'; Action = 'replace'; ImageMso = 'TextBoxInsert'; Image = 'pptk-tk-top' }
    @{ Id = 'PPTKAddBottomTK'; Group = 'Insert'; Label = 'Add Bottom TK'; Kind = 'button'; Action = 'replace'; ImageMso = 'TextBoxInsert'; Image = 'pptk-tk-bottom' }
    @{ Id = 'TextMoreMenu'; Group = 'Text'; Label = 'More'; Kind = 'menu'; Action = 'replace'; ImageMso = 'MenuView2'; Image = 'pptk-text-more' }
    @{ Id = 'SpecialCharacterMenu'; Group = 'Text'; Label = 'Special characters'; Kind = 'menu'; Action = 'replace'; ImageMso = 'SymbolInsert'; Image = 'pptk-special-char' }
    @{ Id = 'PPTKPasteContentOnly'; Group = 'Text'; Label = 'Paste content only'; Kind = 'button'; Action = 'replace'; ImageMso = 'PasteTextOnly'; Image = 'pptk-paste-text' }
    @{ Id = 'PPTKMergeText'; Group = 'Text'; Label = 'Merge textboxes'; Kind = 'button'; Action = 'replace'; ImageMso = 'TextBoxInsert'; Image = 'pptk-merge-text' }
    @{ Id = 'PPTKRemoveMargins'; Group = 'Text'; Label = 'Remove margins'; Kind = 'button'; Action = 'replace'; ImageMso = 'ParagraphMarks'; Image = 'pptk-remove-margins' }
    @{ Id = 'ObjectsAlignLeftsButton'; Group = 'Align'; Label = 'Align left'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectsAlignLeft'; Image = 'pptk-align-left' }
    @{ Id = 'ObjectsAlignCentersButton'; Group = 'Align'; Label = 'Align center'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectsAlignCenterHorizontal'; Image = 'pptk-align-center' }
    @{ Id = 'ObjectsAlignRightsButton'; Group = 'Align'; Label = 'Align right'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectsAlignRight'; Image = 'pptk-align-right' }
    @{ Id = 'ObjectsAlignTopsButton'; Group = 'Align'; Label = 'Align top'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectsAlignTop'; Image = 'pptk-align-top' }
    @{ Id = 'ObjectsAlignMiddlesButton'; Group = 'Align'; Label = 'Align middle'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectsAlignMiddleVertical'; Image = 'pptk-align-middle' }
    @{ Id = 'ObjectsAlignBottomsButton'; Group = 'Align'; Label = 'Align bottom'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectsAlignBottom'; Image = 'pptk-align-bottom' }
    @{ Id = 'ObjectsDistributeHorizontallyButton'; Group = 'Space'; Label = 'Distribute horizontally'; Kind = 'button'; Action = 'replace'; ImageMso = 'AlignDistributeHorizontally'; Image = 'pptk-dist-h' }
    @{ Id = 'ObjectsDistributeVerticallyButton'; Group = 'Space'; Label = 'Distribute vertically'; Kind = 'button'; Action = 'replace'; ImageMso = 'AlignDistributeVertically'; Image = 'pptk-dist-v' }
    @{ Id = 'ObjectsResizeAndSpaceEvenly'; Group = 'Space'; Label = 'Resize and space evenly'; Kind = 'menu'; Action = 'replace'; ImageMso = 'AlignRelativeToPage'; Image = 'pptk-resize-space' }
    @{ Id = 'ResizeAndSpaceEvenHorizontal'; Group = 'Space'; Label = 'Resize and space horizontally (equal spacing)'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-resize-h' }
    @{ Id = 'ResizeAndSpaceEvenHorizontalPreserveFirst'; Group = 'Space'; Label = 'Resize and space horizontally (preserve first)'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-resize-h-first' }
    @{ Id = 'ResizeAndSpaceEvenHorizontalPreserveLast'; Group = 'Space'; Label = 'Resize and space horizontally (preserve last)'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-resize-h-last' }
    @{ Id = 'ResizeAndSpaceEvenVertical'; Group = 'Space'; Label = 'Resize and space vertically (equal spacing)'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-resize-v' }
    @{ Id = 'ResizeAndSpaceEvenVerticalPreserveFirst'; Group = 'Space'; Label = 'Resize and space vertically (preserve first)'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-resize-v-first' }
    @{ Id = 'ResizeAndSpaceEvenVerticalPreserveLast'; Group = 'Space'; Label = 'Resize and space vertically (preserve last)'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-resize-v-last' }
    @{ Id = 'SpaceGapsMenu'; Group = 'Space'; Label = 'Gaps'; Kind = 'menu'; Action = 'replace'; ImageMso = 'AlignDistributeHorizontally'; Image = 'pptk-gaps' }
    @{ Id = 'ObjectsRemoveSpaceHorizontal'; Group = 'Space'; Label = 'Remove horizontal gap between shapes (direction left)'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectsAlignDistributeHorizontallyRemove'; Image = 'pptk-gap-h-left' }
    @{ Id = 'ObjectsRemoveSpaceHorizontalRight'; Group = 'Space'; Label = 'Remove horizontal gap between shapes (direction right)'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectsAlignDistributeHorizontallyRemove'; Image = 'pptk-gap-h-right' }
    @{ Id = 'ObjectsIncreaseSpacingHorizontal'; Group = 'Space'; Label = 'Increase horizontal gap between shapes'; Kind = 'button'; Action = 'replace'; ImageMso = 'HorizontalSpacingIncrease'; Image = 'pptk-gap-h-inc' }
    @{ Id = 'ObjectsDecreaseSpacingHorizontal'; Group = 'Space'; Label = 'Decrease horizontal gap between shapes'; Kind = 'button'; Action = 'replace'; ImageMso = 'HorizontalSpacingDecrease'; Image = 'pptk-gap-h-dec' }
    @{ Id = 'ObjectsRemoveSpaceVertical'; Group = 'Space'; Label = 'Remove vertical gap between shapes (direction up)'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectsAlignDistributeVerticallyRemove'; Image = 'pptk-gap-v-up' }
    @{ Id = 'ObjectsRemoveSpaceVerticalDown'; Group = 'Space'; Label = 'Remove vertical gap between shapes (direction down)'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectsAlignDistributeVerticallyRemove'; Image = 'pptk-gap-v-down' }
    @{ Id = 'ObjectsIncreaseSpacingVertical'; Group = 'Space'; Label = 'Increase vertical gap between shapes'; Kind = 'button'; Action = 'replace'; ImageMso = 'VerticalSpacingIncrease'; Image = 'pptk-gap-v-inc' }
    @{ Id = 'ObjectsDecreaseSpacingVertical'; Group = 'Space'; Label = 'Decrease vertical gap between shapes'; Kind = 'button'; Action = 'replace'; ImageMso = 'VerticalSpacingDecrease'; Image = 'pptk-gap-v-dec' }
    @{ Id = 'PPTKChangeShapeMenu'; Group = 'Shape'; Label = 'Change shape'; Kind = 'menu'; Action = 'replace'; ImageMso = 'ShapeChangeShapeGallery'; Image = 'pptk-change-shape' }
    @{ Id = 'ObjectsSameWidthButton'; Group = 'Shape'; Label = 'Set same width'; Kind = 'button'; Action = 'replace'; ImageMso = 'ColumnWidth'; Image = 'pptk-same-width' }
    @{ Id = 'ObjectsSameHeightButton'; Group = 'Shape'; Label = 'Set same height'; Kind = 'button'; Action = 'replace'; ImageMso = 'RowHeight'; Image = 'pptk-same-height' }
    @{ Id = 'ObjectsSameHeightAndWidthButton'; Group = 'Shape'; Label = 'Set same height and width'; Kind = 'button'; Action = 'replace'; ImageMso = 'SizeToControlHeightAndWidth'; Image = 'pptk-same-size' }
    @{ Id = 'ObjectsSwapPosition'; Group = 'Shape'; Label = 'Swap position of two shapes'; Kind = 'button'; Action = 'replace'; ImageMso = 'CircularReferences'; Image = 'pptk-swap' }
    @{ Id = 'ObjectsSwapPositionMenu'; Group = 'Shape'; Label = 'Swap position of two shapes'; Kind = 'menu'; Action = 'replace'; ImageMso = 'CircularReferences'; Image = 'pptk-swap' }
    @{ Id = 'ObjectsSwapPositionCentered'; Group = 'Shape'; Label = 'Swap position of two shapes (centered)'; Kind = 'button'; Action = 'replace'; ImageMso = 'CircularReferences'; Image = 'pptk-swap-center' }
    @{ Id = 'PPTKSwapText'; Group = 'Shape'; Label = 'Swap text'; Kind = 'button'; Action = 'replace'; ImageMso = 'ReplaceDialog'; Image = 'pptk-swap-text' }
    @{ Id = 'ShapeMoreMenu'; Group = 'Shape'; Label = 'More'; Kind = 'menu'; Action = 'replace'; ImageMso = 'ShapesMoreShapes'; Image = 'pptk-shape-more' }
    @{ Id = 'ObjectsSizeToTallest'; Group = 'Shape'; Label = 'Size to tallest'; Kind = 'button'; Action = 'replace'; ImageMso = 'RowHeight'; Image = 'pptk-size-tall' }
    @{ Id = 'ObjectsSizeShortest'; Group = 'Shape'; Label = 'Size to shortest'; Kind = 'button'; Action = 'replace'; ImageMso = 'RowHeight'; Image = 'pptk-size-short' }
    @{ Id = 'ObjectsSizeToWidest'; Group = 'Shape'; Label = 'Size to widest'; Kind = 'button'; Action = 'replace'; ImageMso = 'ColumnWidth'; Image = 'pptk-size-wide' }
    @{ Id = 'ObjectsSizeNarrowest'; Group = 'Shape'; Label = 'Size to narrowest'; Kind = 'button'; Action = 'replace'; ImageMso = 'ColumnWidth'; Image = 'pptk-size-narrow' }
    @{ Id = 'PPTKAddRightColumn'; Group = 'Shape'; Label = 'Add Right Column'; Kind = 'button'; Action = 'replace'; ImageMso = 'TableInsert'; Image = 'pptk-col-add' }
    @{ Id = 'PPTKAddBottomRow'; Group = 'Shape'; Label = 'Add Bottom Row'; Kind = 'button'; Action = 'replace'; ImageMso = 'TableInsert'; Image = 'pptk-row-add' }
    @{ Id = 'PPTKDeleteRightColumn'; Group = 'Shape'; Label = 'Delete Right Column'; Kind = 'button'; Action = 'replace'; ImageMso = 'TableInsert'; Image = 'pptk-col-del' }
    @{ Id = 'PPTKDeleteBottomRow'; Group = 'Shape'; Label = 'Delete Bottom Row'; Kind = 'button'; Action = 'replace'; ImageMso = 'TableInsert'; Image = 'pptk-row-del' }
    @{ Id = 'ObjectsArrangeShapesButton'; Group = 'Smart Align'; Label = 'Arrange shapes'; Kind = 'button'; Action = 'replace'; ImageMso = 'SlidesPerPage9Slides'; Image = 'pptk-arrange' }
    @{ Id = 'ObjectsStretchMenu'; Group = 'Smart Align'; Label = 'Stretch shapes'; Kind = 'menu'; Action = 'replace'; ImageMso = 'DrawingCanvasExpand'; Image = 'pptk-stretch' }
    @{ Id = 'ObjectsStretchLeft'; Group = 'Smart Align'; Label = 'Stretch shapes to left'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-stretch-left' }
    @{ Id = 'ObjectsStretchRight'; Group = 'Smart Align'; Label = 'Stretch shapes to right'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-stretch-right' }
    @{ Id = 'ObjectsStretchTop'; Group = 'Smart Align'; Label = 'Stretch shapes to top'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-stretch-top' }
    @{ Id = 'ObjectsStretchBottom'; Group = 'Smart Align'; Label = 'Stretch shapes to bottom'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-stretch-bottom' }
    @{ Id = 'ObjectsStretchLeftShapeRight'; Group = 'Smart Align'; Label = 'Stretch shapes to the right edge of the leftmost shape'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-stretch-to-left' }
    @{ Id = 'ObjectsStretchRightShapeLeft'; Group = 'Smart Align'; Label = 'Stretch shapes to the left edge of the rightmost shape'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-stretch-to-right' }
    @{ Id = 'ObjectsStretchTopShapeBottom'; Group = 'Smart Align'; Label = 'Stretch shapes to the bottom edge of the topmost shape'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-stretch-to-top' }
    @{ Id = 'ObjectsStretchBottomShapeTop'; Group = 'Smart Align'; Label = 'Stretch shapes to the top edge of the bottommost shape'; Kind = 'button'; Action = 'add'; ImageMso = ''; Image = 'pptk-stretch-to-bottom' }
    @{ Id = 'SmartAlignMoreMenu'; Group = 'Smart Align'; Label = 'More'; Kind = 'menu'; Action = 'replace'; ImageMso = 'GroupPosition'; Image = 'pptk-align-more' }
    @{ Id = 'PPTKAlignShapeAngle'; Group = 'Smart Align'; Label = 'Align shape angle'; Kind = 'button'; Action = 'replace'; ImageMso = 'ObjectRotateRight90'; Image = 'pptk-align-angle' }
    @{ Id = 'PPTKAlignProcess'; Group = 'Smart Align'; Label = 'Align Process'; Kind = 'button'; Action = 'replace'; ImageMso = 'ShapeChevron'; Image = 'pptk-process' }
    @{ Id = 'PPTKAlignMatrix'; Group = 'Smart Align'; Label = 'Align Matrix'; Kind = 'button'; Action = 'replace'; ImageMso = 'TableInsert'; Image = 'pptk-matrix' }
    @{ Id = 'PPTKAlignCentral'; Group = 'Smart Align'; Label = 'Align Central'; Kind = 'button'; Action = 'replace'; ImageMso = 'GroupPosition'; Image = 'pptk-central' }
    @{ Id = 'PPTKShortcutR'; Group = 'PPTK'; Label = 'Rectangle'; Kind = 'button'; Action = 'replace'; ImageMso = 'ShapeRectangle'; Image = 'pptk-shape-rect' }
    @{ Id = 'PPTKShortcutU'; Group = 'PPTK'; Label = 'Rounded Rectangle'; Kind = 'button'; Action = 'replace'; ImageMso = 'ShapeRoundedRectangle'; Image = 'pptk-shape-round' }
    @{ Id = 'PPTKShortcutO'; Group = 'PPTK'; Label = 'Oval'; Kind = 'button'; Action = 'replace'; ImageMso = 'ShapeOval'; Image = 'pptk-shape-oval' }
    @{ Id = 'PPTKShortcutC'; Group = 'PPTK'; Label = 'Chevron'; Kind = 'button'; Action = 'replace'; ImageMso = 'ShapeChevron'; Image = 'pptk-shape-chevron' }
    @{ Id = 'PPTKShortcutP'; Group = 'PPTK'; Label = 'Pentagon'; Kind = 'button'; Action = 'replace'; ImageMso = 'ShapePentagon'; Image = 'pptk-shape-pentagon' }
    @{ Id = 'PPTKShortcutG'; Group = 'PPTK'; Label = 'Triangle'; Kind = 'button'; Action = 'replace'; ImageMso = 'ShapeIsoscelesTriangle'; Image = 'pptk-shape-triangle' }
    @{ Id = 'PPTKShortcutM'; Group = 'PPTK'; Label = 'Parallelogram'; Kind = 'button'; Action = 'replace'; ImageMso = 'ShapeParallelogram'; Image = 'pptk-shape-para' }
    @{ Id = 'PPTKSmartBar'; Group = 'PPTK'; Label = 'Smart Bar'; Kind = 'button'; Action = 'replace'; ImageMso = 'GroupPosition'; Image = 'pptk-smart-bar' }
    @{ Id = 'ToolsTableMenu'; Group = 'Tools'; Label = 'Table'; Kind = 'menu'; Action = 'replace'; ImageMso = 'TableInsert'; Image = 'pptk-table' }
    @{ Id = 'ToolsMoreMenu'; Group = 'Tools'; Label = 'More Tools'; Kind = 'menu'; Action = 'replace'; ImageMso = 'AdvancedFileProperties'; Image = 'pptk-tools' }
    @{ Id = 'CopyStorylineMenu'; Group = 'Tools'; Label = 'Copy storyline/slide note'; Kind = 'menu'; Action = 'replace'; ImageMso = 'SpeakerNotes'; Image = 'pptk-storyline' }
)

if ($OnlyIcon -eq '') {
$ribbonPath = Join-Path $Root '..\Ribbon\CustomUI.xml'
[xml]$ribbonXml = Get-Content -LiteralPath $ribbonPath -Raw
$live = @{}
foreach ($node in $ribbonXml.GetElementsByTagName('*')) {
    if (-not ($node -is [System.Xml.XmlElement])) { continue }
    if (-not $node.HasAttribute('id')) { continue }
    if ($node.HasAttribute('idMso')) { continue }
    $live[$node.GetAttribute('id')] = [pscustomobject]@{
        ImageMso = $(if ($node.HasAttribute('imageMso')) { $node.GetAttribute('imageMso') } else { '' })
        Name = $node.LocalName
    }
}

$errors = New-Object System.Collections.Generic.List[string]
$seenIds = @{}
foreach ($row in $Map) {
    if ($seenIds.ContainsKey($row.Id)) { $errors.Add("Duplicate manifest id $($row.Id)") }
    $seenIds[$row.Id] = $true
    if (-not $Draw.ContainsKey($row.Image)) { $errors.Add("No artwork for $($row.Image) used by $($row.Id)") }
    if (-not $live.ContainsKey($row.Id)) { $errors.Add("Control $($row.Id) is not in CustomUI.xml") ; continue }
    $actual = $live[$row.Id]
    if ($actual.Name -ne $row.Kind) { $errors.Add("$($row.Id) is $($actual.Name) in XML and $($row.Kind) in the map") }
    if ($actual.ImageMso -ne $row.ImageMso) { $errors.Add("$($row.Id) imageMso is '$($actual.ImageMso)' in XML and '$($row.ImageMso)' in the map") }
}
foreach ($id in $live.Keys) {
    $item = $live[$id]
    if ($item.ImageMso -and -not $seenIds.ContainsKey($id)) {
        $errors.Add("Custom control $id has imageMso $($item.ImageMso) but is missing from the map")
    }
}
if ($errors.Count -gt 0) {
    $errors | ForEach-Object { Write-Output $_ }
    throw "Ribbon map does not match CustomUI.xml."
}
}

function New-IconBitmap([int]$Size, [scriptblock]$Drawer) {
    $bitmap = New-Object System.Drawing.Bitmap $Size, $Size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::None
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::None
    $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceOver
    & $Drawer $graphics ([single]($Size / 16))
    $graphics.Dispose()
    return $bitmap
}

function Save-Png($Bitmap, [string]$Path) {
    $Bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
}

if ($OnlyIcon -ne '') {
    if (-not $Draw.ContainsKey($OnlyIcon)) { throw "Unknown icon $OnlyIcon" }
    $oneSmall = New-IconBitmap 16 $Draw[$OnlyIcon]
    $oneLarge = New-IconBitmap 32 $Draw[$OnlyIcon]
    Save-Png $oneSmall (Join-Path $Icon16 "$OnlyIcon.png")
    Save-Png $oneLarge (Join-Path $Icon32 "$OnlyIcon.png")
    $oneSmall.Dispose()
    $oneLarge.Dispose()
    Write-Output "Generated $OnlyIcon"
    exit 0
}

$built = @{}
foreach ($name in ($Draw.Keys | Sort-Object)) {
    $small = New-IconBitmap 16 $Draw[$name]
    $large = New-IconBitmap 32 $Draw[$name]
    Save-Png $small (Join-Path $Icon16 "$name.png")
    Save-Png $large (Join-Path $Icon32 "$name.png")
    $built[$name] = @{ Small = $small; Large = $large }
}

function Measure-Bitmap($Bitmap) {
    $exact = 0
    $cyanFamily = 0
    $transparent = 0
    $opaque = 0
    $corners = @()
    for ($y = 0; $y -lt $Bitmap.Height; $y++) {
        for ($x = 0; $x -lt $Bitmap.Width; $x++) {
            $pixel = $Bitmap.GetPixel($x, $y)
            if ($pixel.A -eq 0) { $transparent++ } else { $opaque++ }
            if ($pixel.A -gt 200 -and $pixel.R -eq 35 -and $pixel.G -eq 186 -and $pixel.B -eq 236) { $exact++ }
            if ($pixel.A -gt 140 -and $pixel.B -gt ($pixel.R + 30) -and $pixel.G -gt 100 -and $pixel.R -lt 140) { $cyanFamily++ }
        }
    }
    foreach ($point in @(@(0, 0), @(($Bitmap.Width - 1), 0), @(0, ($Bitmap.Height - 1)), @(($Bitmap.Width - 1), ($Bitmap.Height - 1)))) {
        $corners += $Bitmap.GetPixel($point[0], $point[1]).A
    }
    $rect = New-Object System.Drawing.Rectangle 0, 0, $Bitmap.Width, $Bitmap.Height
    $data = $Bitmap.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $byteCount = $data.Stride * $data.Height
    $bytes = New-Object byte[] $byteCount
    [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $bytes, 0, $byteCount)
    $Bitmap.UnlockBits($data)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = [BitConverter]::ToString($sha.ComputeHash($bytes))
    $sha.Dispose()
    return [pscustomobject]@{
        Exact = $exact
        Family = $cyanFamily
        Transparent = $transparent
        Opaque = $opaque
        Corners = ($corners -join ',')
        Hash = $hash
        Format = $Bitmap.PixelFormat.ToString()
    }
}

$report = New-Object System.Collections.Generic.List[string]
$hashes = @{}
$fail = $false
foreach ($name in ($built.Keys | Sort-Object)) {
    foreach ($sizeName in @('Small', 'Large')) {
        $bitmap = $built[$name][$sizeName]
        $expected = $(if ($sizeName -eq 'Small') { 16 } else { 32 })
        $minExact = $(if ($sizeName -eq 'Small') { 8 } else { 24 })
        $stats = Measure-Bitmap $bitmap
        $problems = @()
        if ($bitmap.Width -ne $expected -or $bitmap.Height -ne $expected) { $problems += 'size' }
        if ($stats.Format -ne 'Format32bppArgb') { $problems += 'format' }
        if ($stats.Transparent -lt 8) { $problems += 'transparency' }
        if ($stats.Corners -ne '0,0,0,0') { $problems += "corners=$($stats.Corners)" }
        if ($stats.Exact -lt $minExact) { $problems += "exactCyan=$($stats.Exact)" }
        if ($stats.Family -lt $minExact) { $problems += "cyanFamily=$($stats.Family)" }
        $key = "$sizeName|$($stats.Hash)"
        if ($hashes.ContainsKey($key)) { $problems += "duplicate-of $($hashes[$key])" } else { $hashes[$key] = $name }
        $status = 'OK'
        if ($problems.Count -gt 0) {
            $status = 'FAIL ' + ($problems -join '; ')
            $fail = $true
        }
        $report.Add("$name $expected $($stats.Exact) $($stats.Family) $($stats.Transparent) $($stats.Opaque) $status")
    }
}

function New-Preview([string]$Path, [int]$SourceSize, [int]$Zoom, [System.Drawing.Color]$Back, [string]$Title) {
    $names = @($Draw.Keys | Sort-Object)
    $cols = 7
    $rows = [Math]::Ceiling($names.Count / $cols)
    $cellW = 168
    $cellH = 168
    $header = 36
    $sheet = New-Object System.Drawing.Bitmap ($cols * $cellW), ($header + ($rows * $cellH)), ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($sheet)
    $graphics.Clear($Back)
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $titleBrush = New-Object System.Drawing.SolidBrush $(if ($Back.R -lt 80) { $script:White } else { $script:Ink })
    $font = New-Object System.Drawing.Font 'Segoe UI', 8
    $titleFont = New-Object System.Drawing.Font 'Segoe UI', 11, ([System.Drawing.FontStyle]::Bold)
    $graphics.DrawString($Title, $titleFont, $titleBrush, 8, 8)
    for ($i = 0; $i -lt $names.Count; $i++) {
        $col = $i % $cols
        $row = [Math]::Floor($i / $cols)
        $x = ($col * $cellW) + 16
        $y = $header + ($row * $cellH) + 8
        $graphics.DrawString($names[$i], $font, $titleBrush, ($col * $cellW) + 8, ($y + ($SourceSize * $Zoom) + 6))
        $source = $(if ($SourceSize -eq 16) { $built[$names[$i]].Small } else { $built[$names[$i]].Large })
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::Half
        $dest = New-Object System.Drawing.Rectangle $x, $y, ($SourceSize * $Zoom), ($SourceSize * $Zoom)
        $graphics.DrawImage($source, $dest)
    }
    $font.Dispose()
    $titleFont.Dispose()
    $titleBrush.Dispose()
    $graphics.Dispose()
    Save-Png $sheet $Path
    $sheet.Dispose()
}

New-Preview (Join-Path $Root 'icons\preview-16-light.png') 16 8 $script:Light 'PPTK 16px icons on #F7F9FB'
New-Preview (Join-Path $Root 'icons\preview-16-dark.png') 16 8 $script:Dark 'PPTK 16px icons on #30343B'
New-Preview (Join-Path $Root 'icons\preview-32-light.png') 32 4 $script:Light 'PPTK 32px icons on #F7F9FB'

$brandPath = Join-Path $Root 'ModulePPTKBrand.bas'
$brandText = Get-Content -LiteralPath $brandPath -Raw
function Ole([int]$R, [int]$G, [int]$B) { return $R + ($G * 256) + ($B * 65536) }
$expectedColors = @{
    'PPTK_COLOR_PRIMARY' = (Ole 35 186 236)
    'PPTK_COLOR_WHITE' = (Ole 255 255 255)
    'PPTK_COLOR_DARK_TEXT' = (Ole 48 52 59)
    'PPTK_COLOR_LIGHT_BACKGROUND' = (Ole 247 249 251)
    'PPTK_COLOR_SUBTLE_BORDER' = (Ole 227 232 237)
}
foreach ($constName in $expectedColors.Keys) {
    $match = [regex]::Match($brandText, "Public Const $constName As Long = (\d+)")
    if (-not $match.Success) { $report.Add("BRAND FAIL missing $constName"); $fail = $true; continue }
    $actualValue = [int]$match.Groups[1].Value
    if ($actualValue -ne $expectedColors[$constName]) {
        $report.Add("BRAND FAIL $constName is $actualValue expected $($expectedColors[$constName])")
        $fail = $true
    } else {
        $report.Add("BRAND OK $constName = $actualValue")
    }
}

$uniqueImages = @($Map | ForEach-Object { $_.Image } | Select-Object -Unique | Sort-Object)
$rels = New-Object System.Text.StringBuilder
[void]$rels.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$rels.AppendLine('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
foreach ($image in $uniqueImages) {
    [void]$rels.AppendLine("  <Relationship Id=""$image"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""images/$image.png""/>")
}
[void]$rels.AppendLine('</Relationships>')
[System.IO.File]::WriteAllText((Join-Path $Root 'customUI.xml.rels'), $rels.ToString().Replace("`r`n", "`n").Replace("`n", "`r`n"))

$md = New-Object System.Text.StringBuilder
[void]$md.AppendLine('# PPTK cyan ribbon integration manifest')
[void]$md.AppendLine('')
[void]$md.AppendLine('These 16px icons are installed in `src/PPTK/Ribbon/CustomUI.xml` and in `src/PPTK-WINDOWS-FINAL.pptm`. The source ribbon and the embedded ribbon are the same file.')
[void]$md.AppendLine('')
[void]$md.AppendLine('Apply this only after the Smart Bar, Align Matrix, and Align Central UI task has finished editing the ribbon and the PPTM.')
[void]$md.AppendLine('')
[void]$md.AppendLine('## Brand colors')
[void]$md.AppendLine('')
[void]$md.AppendLine('| Role | Hex | VBA RGB | OLE Long |')
[void]$md.AppendLine('| --- | --- | --- | ---: |')
[void]$md.AppendLine('| Primary cyan | #23BAEC | RGB(35, 186, 236) | 15514147 |')
[void]$md.AppendLine('| White | #FFFFFF | RGB(255, 255, 255) | 16777215 |')
[void]$md.AppendLine('| Dark text | #30343B | RGB(48, 52, 59) | 3879984 |')
[void]$md.AppendLine('| Light background | #F7F9FB | RGB(247, 249, 251) | 16513527 |')
[void]$md.AppendLine('| Subtle border | #E3E8ED | RGB(227, 232, 237) | 15591651 |')
[void]$md.AppendLine('')
[void]$md.AppendLine('UserForm helpers live in `ModulePPTKBrand.bas`. That module is imported in the Windows VBA project. `PPTKSmartBarForm`, `PPTKMatrixForm`, and `PPTKCentralForm` call it. Do not import it into the Mac workbook.')
[void]$md.AppendLine('')
[void]$md.AppendLine('| Helper | Result |')
[void]$md.AppendLine('| --- | --- |')
[void]$md.AppendLine('| PPTK_StyleForm | White background, dark text |')
[void]$md.AppendLine('| PPTK_StylePrimaryButton | Cyan background, white text, no border |')
[void]$md.AppendLine('| PPTK_StyleSecondaryButton | Light background, dark text, subtle border |')
[void]$md.AppendLine('| PPTK_StyleField | White background, dark text, subtle border |')
[void]$md.AppendLine('')
[void]$md.AppendLine('## Icon files')
[void]$md.AppendLine('')
[void]$md.AppendLine('Embed the 16px PNGs from `icons/16`. The current ribbon uses `size="normal"`, which is a 16px image in the Office 2006 custom UI schema. The 32px PNGs in `icons/32` are the same artwork at 2x for inspection or a later large-button test. Do not embed a 32px file on a normal button; Office will scale it down.')
[void]$md.AppendLine('')
[void]$md.AppendLine('Backgrounds are transparent. Cyan is #23BAEC. Ink is #30343B. White is used only as a cutout.')
[void]$md.AppendLine('')
[void]$md.AppendLine('## Control mapping')
[void]$md.AppendLine('')
[void]$md.AppendLine('Edit each control by its `id`. Do not search-and-replace an `imageMso` value: several controls share one Office icon today and need different cyan icons.')
[void]$md.AppendLine('')
[void]$md.AppendLine('`replace` removes that control''s `imageMso` attribute and adds `image`. `add` only adds `image` on a custom control that currently has no icon. Leave `id`, `label`, `onAction`, `tag`, `size`, `screentip`, `supertip`, and group order unchanged.')
[void]$md.AppendLine('')
[void]$md.AppendLine('| Group | Control id | Action | Remove imageMso | image value | 16px file |')
[void]$md.AppendLine('| --- | --- | --- | --- | --- | --- |')
foreach ($row in $Map) {
    $remove = $(if ($row.Action -eq 'replace') { $row.ImageMso } else { '' })
    [void]$md.AppendLine("| $($row.Group) | ``$($row.Id)`` | $($row.Action) | $remove | ``$($row.Image)`` | icons/16/$($row.Image).png |")
}
[void]$md.AppendLine('')
[void]$md.AppendLine('`ObjectsSwapPosition` and `ObjectsSwapPositionMenu` share `pptk-swap` because they are the same swap command. Every other image is used by one control.')
[void]$md.AppendLine('')
[void]$md.AppendLine('## Controls that stay on Office icons')
[void]$md.AppendLine('')
[void]$md.AppendLine('Any control with `idMso` keeps its built-in Office icon. The 2006 ribbon schema does not accept a custom image on an `idMso` control, and replacing those controls would change command behavior. That includes fill, outline, font, paste PNG, native align toggles, and the other built-in buttons already on the PPTK tab.')
[void]$md.AppendLine('')
[void]$md.AppendLine('## How to integrate')
[void]$md.AppendLine('')
[void]$md.AppendLine('1. Wait until no other task is editing `src/PPTK/Ribbon/CustomUI.xml` or `src/PPTK-WINDOWS-FINAL.pptm`.')
[void]$md.AppendLine('2. Re-run `Generate-PptkRibbonIcons.ps1`. It fails if a mapped control, element type, or `imageMso` no longer matches the XML.')
[void]$md.AppendLine('3. In the then-current `CustomUI.xml`, apply the table above one control at a time. Do not replace the whole file from an older copy.')
[void]$md.AppendLine('4. Copy `icons/16/*.png` into the PPTM part `customUI/images/`.')
[void]$md.AppendLine('5. The package relationship `PPTKRibbon` already points at `customUI/customUI.xml`. Leave `_rels/.rels` unchanged.')
[void]$md.AppendLine('6. `customUI/_rels/customUI.xml.rels` does not exist in the PPTM today. If it is still missing, add `customUI.xml.rels` from this folder. If the UI task created that file, insert only Relationship rows whose Id is not already present.')
[void]$md.AppendLine('7. `[Content_Types].xml` already has a PNG default. Do not add another PNG content type.')
[void]$md.AppendLine('8. Import `ModulePPTKBrand.bas` with the VBA editor when a UserForm is ready to call it. Do not call it from Smart Bar, Align Matrix, Align Central, shortcut dispatch, or keyboard code as part of the icon install.')
[void]$md.AppendLine('9. Reload the add-in and look at the PPTK tab at normal ribbon size. The icons are not live until this step has been done.')
[void]$md.AppendLine('')
[void]$md.AppendLine('Office RibbonX Editor can do steps 4 and 6 if the `image` values match the relationship Ids in `customUI.xml.rels`.')
[void]$md.AppendLine('')
$manifestPath = Join-Path $Root 'ICON-MANIFEST.md'
[System.IO.File]::WriteAllText($manifestPath, $md.ToString().Replace("`r`n", "`n").Replace("`n", "`r`n"))

$reportPath = Join-Path $Root 'icon-validation.txt'
$header = @(
    "icons=$($built.Count)"
    "mappedControls=$($Map.Count)"
    "uniqueImages=$($uniqueImages.Count)"
    'columns: name size exactCyan cyanFamily transparent opaque status'
)
[System.IO.File]::WriteAllText($reportPath, (($header + $report) -join "`r`n"))

foreach ($name in $built.Keys) {
    $built[$name].Small.Dispose()
    $built[$name].Large.Dispose()
}

if ($fail) { throw 'Icon or brand validation failed. See icon-validation.txt.' }
Write-Output "Generated $($built.Count) icons for $($Map.Count) controls."
