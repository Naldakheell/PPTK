Attribute VB_Name = "ModulePPTKFixColors"
Option Explicit

' PPTK Fix Slide Colors (Q). This is a PPTK mapping, not auxi parity.
'
' Scope:
'   The slide currently shown in Normal view. Other slides, layouts, and
'   masters are not modified. Reading the slide's theme does not write it.
'
' Recolored, when the color is an absolute RGB color:
'   Solid fills, solid outlines, gradient stops, font runs,
'   table cell fills, table borders, table text,
'   chart series and point solid fills and lines,
'   SmartArt node solid fills and text when the object model allows it,
'   and a slide's own solid background when it does not follow the master.
'   Grouped shapes are visited through their descendants.
'
' Left unchanged:
'   Pictures, linked pictures, media, OLE objects, and ink.
'   Shapes whose name contains "logo".
'   Picture, texture, and pattern fills.
'   Fills and lines that are hidden or at least 98% transparent.
'   Colors that are already theme scheme colors.
'   Mid-neutral grays (chroma <= 0.12 and luminance between 0.28 and 0.72).
'   Shadows, glows, and any object the object model refuses to change.
'
' Mapping, matching ThemeHarmonize.cs:
'   Near-white luminance >= 0.90 links to Light 1.
'   Near-black luminance <= 0.12 links to Dark 1.
'   Other light neutrals link to Light 2, other dark neutrals to Dark 2.
'   Chromatic colors link to the accent with the nearest hue.
'   An equal hue distance keeps the earlier accent.
'   Neutral accents are not hue targets. Brightness is the luminance
'   difference, clamped to +/- 0.55, so light and dark shades stay apart.
'   After a text color is applied, a contrast ratio below 3 against the
'   surface behind the text replaces that text with Dark 1 on a light surface
'   or Light 1 on a dark surface. The surface is the shape's own solid fill;
'   for a shape with no fill, the topmost lower shape under the text's center
'   with a solid fill (fully transparent fills are looked through), else the
'   slide's own solid background. A picture, gradient, pattern, chart, table,
'   SmartArt, or semi-transparent fill behind the text, or a non-solid slide
'   background, is unknown: the text keeps its theme mapping and no contrast
'   color is forced.
'
' Undo:
'   The procedure makes one pass of object-model writes. PowerPoint does not
'   provide a VBA transaction around those writes. Whether they undo as one
'   step has to be checked in PowerPoint and is not claimed here.

Private Const NeutralChroma As Double = 0.12
Private Const WhiteLuminance As Double = 0.9
Private Const BlackLuminance As Double = 0.12
Private Const LightNeutralLuminance As Double = 0.72
Private Const DarkNeutralLuminance As Double = 0.28
Private Const MaxBrightness As Double = 0.55
Private Const MinimumContrast As Double = 3#
Private Const InvisibleTransparency As Single = 0.98
Private Const SeeThroughTransparency As Single = 0.5
Private Const UnknownLuminance As Double = -1#

Private PalDark1 As Long
Private PalLight1 As Long
Private PalDark2 As Long
Private PalLight2 As Long
Private PalAccent(1 To 6) As Long
Private PageLuminance As Double

Public Sub PPTK_FixSlideColors()
    Dim sld As Slide
    Dim i As Long

    On Error GoTo Clean
    If Not PPTK_CurrentSlide(sld) Then Exit Sub
    If Not PPTK_LoadPalette(sld) Then Exit Sub
    PageLuminance = PPTK_SlideBackgroundLuminance(sld)
    PPTK_RecolorOwnBackground sld

    For i = 1 To sld.Shapes.Count
        PPTK_RecolorShape sld.Shapes(i)
    Next i
Clean:
End Sub

Private Function PPTK_CurrentSlide(ByRef sld As Slide) As Boolean
    PPTK_CurrentSlide = False
    Set sld = Nothing
    On Error GoTo Clean
    If Application.Presentations.Count < 1 Then Exit Function
    If ActiveWindow Is Nothing Then Exit Function
    If ActiveWindow.ViewType <> ppViewNormal Then Exit Function
    Set sld = ActiveWindow.View.Slide
    PPTK_CurrentSlide = Not sld Is Nothing
Clean:
End Function

Private Function PPTK_LoadPalette(ByVal sld As Slide) As Boolean
    On Error GoTo Clean
    PPTK_ReadScheme sld.Design.SlideMaster.Theme.ThemeColorScheme
    PPTK_LoadPalette = True
    Exit Function
Clean:
    On Error GoTo Failed
    PPTK_ReadScheme ActivePresentation.SlideMaster.Theme.ThemeColorScheme
    PPTK_LoadPalette = True
    Exit Function
Failed:
    PPTK_LoadPalette = False
End Function

Private Sub PPTK_ReadScheme(ByVal scheme As Object)
    PalDark1 = scheme(msoThemeDark1).RGB
    PalLight1 = scheme(msoThemeLight1).RGB
    PalDark2 = scheme(msoThemeDark2).RGB
    PalLight2 = scheme(msoThemeLight2).RGB
    PalAccent(1) = scheme(msoThemeAccent1).RGB
    PalAccent(2) = scheme(msoThemeAccent2).RGB
    PalAccent(3) = scheme(msoThemeAccent3).RGB
    PalAccent(4) = scheme(msoThemeAccent4).RGB
    PalAccent(5) = scheme(msoThemeAccent5).RGB
    PalAccent(6) = scheme(msoThemeAccent6).RGB
End Sub

' The slide's effective background when it is a visible solid fill; otherwise unknown.
Private Function PPTK_SlideBackgroundLuminance(ByVal sld As Slide) As Double
    PPTK_SlideBackgroundLuminance = UnknownLuminance
    On Error GoTo Clean
    If sld.Background.Fill.Visible <> msoTrue Then Exit Function
    If sld.Background.Fill.Type <> msoFillSolid Then Exit Function
    PPTK_SlideBackgroundLuminance = PPTK_Luminance(sld.Background.Fill.ForeColor.RGB)
Clean:
End Function

Private Sub PPTK_RecolorOwnBackground(ByVal sld As Slide)
    On Error Resume Next
    If sld.FollowMasterBackground Then Exit Sub
    If sld.Background.Fill.Type <> msoFillSolid Then Exit Sub
    If Not PPTK_PaintFormat(sld.Background.Fill.ForeColor) Then Exit Sub
    PageLuminance = PPTK_Luminance(sld.Background.Fill.ForeColor.RGB)
End Sub

Private Sub PPTK_RecolorShape(ByVal shp As Shape)
    On Error GoTo Clean
    If shp Is Nothing Then Exit Sub
    If PPTK_IsImageLike(shp) Then Exit Sub
    If InStr(1, shp.Name, "logo", vbTextCompare) > 0 Then Exit Sub

    If shp.Type = msoGroup Then
        Dim child As Shape
        For Each child In shp.GroupItems
            PPTK_RecolorShape child
        Next child
        Exit Sub
    End If

    If shp.Type = msoCanvas Then
        Dim canvasShape As Shape
        For Each canvasShape In shp.CanvasItems
            PPTK_RecolorShape canvasShape
        Next canvasShape
        Exit Sub
    End If

    If shp.HasTable Then
        PPTK_RecolorTable shp.Table, PPTK_BehindLuminance(shp)
        Exit Sub
    End If

    If shp.HasChart Then
        PPTK_RecolorChart shp.Chart
        PPTK_RecolorText shp, PageLuminance
        Exit Sub
    End If

    If shp.HasSmartArt Then
        PPTK_RecolorSmartArt shp
    End If

    Dim fillLuminance As Double
    fillLuminance = PPTK_BehindLuminance(shp)
    PPTK_RecolorFill shp, fillLuminance
    PPTK_RecolorLine shp
    PPTK_RecolorText shp, fillLuminance
Clean:
End Sub

Private Function PPTK_IsImageLike(ByVal shp As Shape) As Boolean
    Select Case shp.Type
        Case msoPicture, msoLinkedPicture, msoMedia, msoOLEControlObject, _
             msoEmbeddedOLEObject, msoLinkedOLEObject, msoInk, msoInkComment
            PPTK_IsImageLike = True
        Case Else
            PPTK_IsImageLike = False
    End Select
End Function

Private Sub PPTK_RecolorFill(ByVal shp As Shape, ByRef fillLuminance As Double)
    On Error GoTo Clean
    If shp.Fill.Visible <> msoTrue Then Exit Sub
    If shp.Fill.Transparency >= InvisibleTransparency Then Exit Sub
    Select Case shp.Fill.Type
        Case msoFillSolid
            ' The text sits on this fill whether or not the fill was recolored
            ' (mid-grays and scheme colors are left as they are).
            PPTK_PaintFormat shp.Fill.ForeColor
            fillLuminance = PPTK_Luminance(shp.Fill.ForeColor.RGB)
        Case msoFillGradient
            Dim stopIndex As Long
            For stopIndex = 1 To shp.Fill.GradientStops.Count
                PPTK_PaintFormat shp.Fill.GradientStops(stopIndex).Color
            Next stopIndex
            fillLuminance = UnknownLuminance
        Case Else
            fillLuminance = UnknownLuminance
    End Select
Clean:
End Sub

Private Sub PPTK_RecolorLine(ByVal shp As Shape)
    On Error GoTo Clean
    If shp.Line.Visible <> msoTrue Then Exit Sub
    If shp.Line.Transparency >= InvisibleTransparency Then Exit Sub
    PPTK_PaintFormat shp.Line.ForeColor
Clean:
End Sub

Private Sub PPTK_RecolorText(ByVal shp As Shape, ByVal backgroundLuminance As Double)
    Dim rng As TextRange
    Dim index As Long
    Dim runStart As Long
    Dim runRgb As Long
    Dim length As Long

    On Error GoTo Clean
    If Not shp.HasTextFrame Then Exit Sub
    If Not shp.TextFrame.HasText Then Exit Sub
    Set rng = shp.TextFrame.TextRange
    length = rng.Length
    If length < 1 Then Exit Sub

    runStart = 1
    runRgb = rng.Characters(1, 1).Font.Color.RGB
    For index = 2 To length + 1
        Dim current As Long
        current = runRgb
        If index <= length Then current = rng.Characters(index, 1).Font.Color.RGB
        If index > length Or current <> runRgb Then
            PPTK_PaintTextRun rng, runStart, index - runStart, backgroundLuminance
            runStart = index
            runRgb = current
        End If
    Next index
Clean:
End Sub

Private Sub PPTK_PaintTextRun(ByVal rng As TextRange, ByVal startAt As Long, ByVal count As Long, ByVal backgroundLuminance As Double)
    Dim fmt As ColorFormat
    On Error GoTo Clean
    If count < 1 Then Exit Sub
    Set fmt = rng.Characters(startAt, count).Font.Color
    If Not PPTK_PaintFormat(fmt) Then Exit Sub
    If backgroundLuminance < 0 Then Exit Sub
    If ThemeNeedsContrast(PPTK_Luminance(fmt.RGB), backgroundLuminance) Then
        If backgroundLuminance >= 0.5 Then
            fmt.ObjectThemeColor = msoThemeColorDark1
        Else
            fmt.ObjectThemeColor = msoThemeColorLight1
        End If
        fmt.Brightness = 0
    End If
Clean:
End Sub

Private Function ThemeNeedsContrast(ByVal textLuminance As Double, ByVal backgroundLuminance As Double) As Boolean
    Dim lighter As Double
    Dim darker As Double
    lighter = textLuminance
    darker = backgroundLuminance
    If darker > lighter Then
        lighter = backgroundLuminance
        darker = textLuminance
    End If
    ThemeNeedsContrast = (lighter + 0.05) / (darker + 0.05) < MinimumContrast
End Function

Private Sub PPTK_RecolorTable(ByVal tbl As Table, ByVal behindLuminance As Double)
    Dim r As Long
    Dim c As Long
    Dim cellShape As Shape
    Dim fillLuminance As Double

    On Error GoTo Clean
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            On Error Resume Next
            Set cellShape = tbl.Cell(r, c).Shape
            If Not cellShape Is Nothing Then
                fillLuminance = behindLuminance
                PPTK_RecolorFill cellShape, fillLuminance
                PPTK_RecolorCellBorders tbl.Cell(r, c)
                PPTK_RecolorText cellShape, fillLuminance
            End If
            On Error GoTo Clean
        Next c
    Next r
Clean:
End Sub

Private Sub PPTK_RecolorCellBorders(ByVal cell As Cell)
    On Error Resume Next
    PPTK_PaintBorder cell.Borders(ppBorderTop)
    PPTK_PaintBorder cell.Borders(ppBorderBottom)
    PPTK_PaintBorder cell.Borders(ppBorderLeft)
    PPTK_PaintBorder cell.Borders(ppBorderRight)
End Sub

Private Sub PPTK_PaintBorder(ByVal border As LineFormat)
    On Error GoTo Clean
    If border.Visible <> msoTrue Then Exit Sub
    PPTK_PaintFormat border.ForeColor
Clean:
End Sub

Private Sub PPTK_RecolorChart(ByVal cht As Chart)
    Dim seriesIndex As Long
    Dim pointIndex As Long

    On Error GoTo Clean
    For seriesIndex = 1 To cht.SeriesCollection.Count
        On Error Resume Next
        PPTK_PaintFormat cht.SeriesCollection(seriesIndex).Format.Fill.ForeColor
        PPTK_PaintFormat cht.SeriesCollection(seriesIndex).Format.Line.ForeColor
        For pointIndex = 1 To cht.SeriesCollection(seriesIndex).Points.Count
            PPTK_PaintFormat cht.SeriesCollection(seriesIndex).Points(pointIndex).Format.Fill.ForeColor
            PPTK_PaintFormat cht.SeriesCollection(seriesIndex).Points(pointIndex).Format.Line.ForeColor
        Next pointIndex
        On Error GoTo Clean
    Next seriesIndex
Clean:
End Sub

Private Sub PPTK_RecolorSmartArt(ByVal shp As Shape)
    Dim nodeIndex As Long
    Dim nodeShape As Shape
    Dim fillLuminance As Double

    On Error GoTo Clean
    For nodeIndex = 1 To shp.SmartArt.AllNodes.Count
        On Error Resume Next
        For Each nodeShape In shp.SmartArt.AllNodes(nodeIndex).Shapes
            fillLuminance = PageLuminance
            PPTK_RecolorFill nodeShape, fillLuminance
            PPTK_RecolorLine nodeShape
            PPTK_RecolorText nodeShape, fillLuminance
        Next nodeShape
        On Error GoTo Clean
    Next nodeIndex
Clean:
End Sub

' Luminance of what is behind a shape that shows no fill of its own: the topmost
' lower shape whose box contains the shape's center (earlier group siblings first,
' then top-level shapes below), else the page. Unknown (-1) when that surface is
' a picture, gradient, pattern, chart, table, SmartArt, or semi-transparent fill.
Private Function PPTK_BehindLuminance(ByVal shp As Shape) As Double
    Dim cx As Single
    Dim cy As Single
    Dim topShape As Shape
    Dim sld As Object
    Dim index As Long
    Dim candidate As Shape
    Dim lum As Double
    Dim state As Long

    PPTK_BehindLuminance = UnknownLuminance
    On Error GoTo Clean
    cx = shp.Left + shp.Width / 2
    cy = shp.Top + shp.Height / 2
    Set topShape = shp
    If shp.Child = msoTrue Then
        Set topShape = shp.ParentGroup
        For index = topShape.GroupItems.Count To 1 Step -1
            Set candidate = topShape.GroupItems(index)
            If candidate.ZOrderPosition < shp.ZOrderPosition Then
                state = PPTK_SurfaceAt(candidate, cx, cy, lum)
                If state = 1 Then PPTK_BehindLuminance = lum: Exit Function
                If state = -1 Then Exit Function
            End If
        Next index
    End If
    Set sld = topShape.Parent
    For index = sld.Shapes.Count To 1 Step -1
        Set candidate = sld.Shapes(index)
        If candidate.ZOrderPosition < topShape.ZOrderPosition Then
            state = PPTK_SurfaceAt(candidate, cx, cy, lum)
            If state = 1 Then PPTK_BehindLuminance = lum: Exit Function
            If state = -1 Then Exit Function
        End If
    Next index
    PPTK_BehindLuminance = PageLuminance
Clean:
End Function

' 1 = solid surface (lum set), 0 = not under the point or see-through, -1 = unknown surface.
Private Function PPTK_SurfaceAt(ByVal shp As Shape, ByVal x As Single, ByVal y As Single, ByRef lum As Double) As Long
    Dim index As Long
    Dim state As Long
    PPTK_SurfaceAt = 0
    On Error GoTo Unknown
    If shp.Visible <> msoTrue Then Exit Function
    If x < shp.Left Or x > shp.Left + shp.Width Or y < shp.Top Or y > shp.Top + shp.Height Then Exit Function
    If shp.Type = msoGroup Then
        For index = shp.GroupItems.Count To 1 Step -1
            state = PPTK_SurfaceAt(shp.GroupItems(index), x, y, lum)
            If state <> 0 Then
                PPTK_SurfaceAt = state
                Exit Function
            End If
        Next index
        Exit Function
    End If
    If PPTK_IsImageLike(shp) Or shp.HasChart Or shp.HasTable Or shp.HasSmartArt Then GoTo Unknown
    If shp.Fill.Visible <> msoTrue Then Exit Function
    If shp.Fill.Transparency >= InvisibleTransparency Then Exit Function
    If shp.Fill.Type <> msoFillSolid Then GoTo Unknown
    If shp.Fill.Transparency > SeeThroughTransparency Then GoTo Unknown
    lum = PPTK_Luminance(shp.Fill.ForeColor.RGB)
    PPTK_SurfaceAt = 1
    Exit Function
Unknown:
    PPTK_SurfaceAt = -1
End Function

Private Function PPTK_PaintFormat(ByVal fmt As ColorFormat) As Boolean
    Dim slot As MsoThemeColorIndex
    Dim brightness As Single
    On Error GoTo Clean
    PPTK_PaintFormat = False
    If fmt Is Nothing Then Exit Function
    If fmt.Type = msoColorTypeScheme Then Exit Function
    If Not PPTK_MapRgb(fmt.RGB, slot, brightness) Then Exit Function
    fmt.ObjectThemeColor = slot
    fmt.Brightness = brightness
    PPTK_PaintFormat = True
Clean:
End Function

Private Function PPTK_MapRgb(ByVal value As Long, ByRef slot As MsoThemeColorIndex, ByRef brightness As Single) As Boolean
    Dim sourceL As Double
    Dim chroma As Double
    Dim bestDistance As Double
    Dim distance As Double
    Dim index As Long
    Dim chosen As Long
    Dim found As Boolean

    sourceL = PPTK_Luminance(value)
    chroma = PPTK_Chroma(value)
    brightness = 0
    PPTK_MapRgb = False

    If chroma <= NeutralChroma Then
        If sourceL >= WhiteLuminance Then
            slot = msoThemeColorLight1
            PPTK_MapRgb = True
        ElseIf sourceL <= BlackLuminance Then
            slot = msoThemeColorDark1
            PPTK_MapRgb = True
        ElseIf sourceL >= LightNeutralLuminance Then
            slot = msoThemeColorLight2
            brightness = PPTK_ClampBrightness(sourceL - PPTK_Luminance(PalLight2))
            PPTK_MapRgb = True
        ElseIf sourceL <= DarkNeutralLuminance Then
            slot = msoThemeColorDark2
            brightness = PPTK_ClampBrightness(sourceL - PPTK_Luminance(PalDark2))
            PPTK_MapRgb = True
        End If
        Exit Function
    End If

    bestDistance = 9999
    chosen = 1
    found = False
    For index = 1 To 6
        If PPTK_Chroma(PalAccent(index)) > NeutralChroma Then
            distance = PPTK_HueDistance(PPTK_Hue(value), PPTK_Hue(PalAccent(index)))
            If Not found Or distance < bestDistance - 0.000000001 Then
                found = True
                bestDistance = distance
                chosen = index
            End If
        End If
    Next index

    slot = msoThemeColorAccent1 + (chosen - 1)
    brightness = PPTK_ClampBrightness(sourceL - PPTK_Luminance(PalAccent(chosen)))
    PPTK_MapRgb = True
End Function

Private Function PPTK_ClampBrightness(ByVal value As Double) As Single
    If value > MaxBrightness Then value = MaxBrightness
    If value < -MaxBrightness Then value = -MaxBrightness
    PPTK_ClampBrightness = CSng(value)
End Function

Private Function PPTK_Luminance(ByVal value As Long) As Double
    PPTK_Luminance = 0.2126 * PPTK_Channel(PPTK_Red(value)) _
        + 0.7152 * PPTK_Channel(PPTK_Green(value)) _
        + 0.0722 * PPTK_Channel(PPTK_Blue(value))
End Function

Private Function PPTK_Channel(ByVal component As Long) As Double
    Dim c As Double
    c = component / 255#
    If c <= 0.04045 Then
        PPTK_Channel = c / 12.92
    Else
        PPTK_Channel = ((c + 0.055) / 1.055) ^ 2.4
    End If
End Function

Private Function PPTK_Chroma(ByVal value As Long) As Double
    Dim r As Double
    Dim g As Double
    Dim b As Double
    Dim maxC As Double
    Dim minC As Double
    r = PPTK_Red(value) / 255#
    g = PPTK_Green(value) / 255#
    b = PPTK_Blue(value) / 255#
    maxC = r
    If g > maxC Then maxC = g
    If b > maxC Then maxC = b
    minC = r
    If g < minC Then minC = g
    If b < minC Then minC = b
    PPTK_Chroma = maxC - minC
End Function

Private Function PPTK_Hue(ByVal value As Long) As Double
    Dim r As Double
    Dim g As Double
    Dim b As Double
    Dim maxC As Double
    Dim minC As Double
    Dim d As Double
    Dim h As Double
    r = PPTK_Red(value) / 255#
    g = PPTK_Green(value) / 255#
    b = PPTK_Blue(value) / 255#
    maxC = r
    If g > maxC Then maxC = g
    If b > maxC Then maxC = b
    minC = r
    If g < minC Then minC = g
    If b < minC Then minC = b
    d = maxC - minC
    If d < 0.00000001 Then
        PPTK_Hue = 0
        Exit Function
    End If
    If maxC = r Then
        h = (g - b) / d
    ElseIf maxC = g Then
        h = (b - r) / d + 2
    Else
        h = (r - g) / d + 4
    End If
    If h < 0 Then h = h + 6
    PPTK_Hue = h * 60#
End Function

Private Function PPTK_HueDistance(ByVal a As Double, ByVal b As Double) As Double
    Dim d As Double
    d = Abs(a - b)
    Do While d >= 360#
        d = d - 360#
    Loop
    If d > 180# Then d = 360# - d
    PPTK_HueDistance = d
End Function

Private Function PPTK_Red(ByVal value As Long) As Long
    PPTK_Red = value And 255
End Function

Private Function PPTK_Green(ByVal value As Long) As Long
    PPTK_Green = (value \ 256) And 255
End Function

Private Function PPTK_Blue(ByVal value As Long) As Long
    PPTK_Blue = (value \ 65536) And 255
End Function
