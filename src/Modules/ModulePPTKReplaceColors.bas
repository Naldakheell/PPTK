Attribute VB_Name = "ModulePPTKReplaceColors"
Option Explicit

' PPTK Find & Replace Colors.
'
' Scope:
'   Ordinary slides only, by SlideIndex: all slides, the current slide, or a
'   From/To range. Slide masters, layouts, notes pages and handouts are never
'   read or written. Slide backgrounds are not a color type here.
'
' Color types (each optional):
'   Fill    - a visible solid shape fill, and a native table cell's solid fill.
'   Outline - a visible shape outline (line). Table borders are not included.
'   Font    - text color. A run of text with one color counts once per shape:
'             adjacent runs with the same color are one use, so a bold word in
'             the middle of black text does not add a second black use.
'   One use is one property of one object: a shape's fill, a shape's outline,
'   or one same-colored span of a shape's text. Nothing is counted twice.
'
' Visited:
'   AutoShapes, text boxes, placeholders, freeforms, lines and connectors (outline),
'   groups (every child, at any depth), native tables (cell fill and cell text).
'
' Skipped and reported, never changed:
'   Pictures, media, OLE and ActiveX objects, ink, charts, SmartArt.
'   Gradient, pattern, texture and picture fills.
'   A property the object model refuses to read.
'
' Matching is exact RGB on the color PowerPoint displays. ColorFormat.RGB returns
' the displayed color for explicit RGB, for theme colors, for theme colors with a
' Lighter/Darker step (Brightness) and for shape-style tints and shades
' (TintAndShade); this was checked against rendered slide pixels. Replacing writes
' ForeColor.RGB, which makes that one property an explicit RGB color and clears its
' tint/shade; the theme itself and other objects using the theme color are not
' changed. Transparency, weight, dash, font, size, text and geometry are kept.
'
' Safety: before Apply changes anything, a copy of the presentation is saved to
' Documents\PPTK Color Backups (see PPTK_SaveBackupCopy). All writes follow one
' StartNewUndoEntry.

Private Const KindFill As Long = 1
Private Const KindLine As Long = 2
Private Const KindFont As Long = 3

' Scan results
Private mCount As Long
Private mRgb() As Long
Private mFill() As Long
Private mLine() As Long
Private mFont() As Long
Private mSlides() As String          ' ",1,4,7," for fast membership

' Walk settings
Private mDoFill As Boolean
Private mDoLine As Boolean
Private mDoFont As Boolean
Private mReplace As Boolean
Private mSources As String           ' ",123,456," RGB values to replace
Private mTarget As Long
Private mSlideIndex As Long

' Replace results and skip notes
Private mChanged As Long
Private mFailed As Long
Private mWriteFailed As Long
Private mChangedSlides As String
Private mSkipPictures As Long
Private mSkipCharts As Long
Private mSkipSmartArt As Long
Private mSkipObjects As Long
Private mSkipEffects As Long

Public Sub PPTK_ShowReplaceColors()
    On Error GoTo Failed
    If Application.Presentations.Count < 1 Then
        MsgBox "Open a presentation first.", vbInformation, "PPTK"
        Exit Sub
    End If
    If ActivePresentation.Slides.Count < 1 Then
        MsgBox "The presentation has no slides.", vbInformation, "PPTK"
        Exit Sub
    End If
    PPTKColorsForm.Show vbModal
    Exit Sub
Failed:
    MsgBox "Find & Replace Colors could not open: " & Err.Description, vbExclamation, "PPTK"
End Sub

' Current slide in Normal view, else the first selected slide; 0 when unknown.
Public Function PPTK_ColorCurrentSlide() As Long
    On Error Resume Next
    PPTK_ColorCurrentSlide = 0
    PPTK_ColorCurrentSlide = ActiveWindow.View.Slide.SlideIndex
    If PPTK_ColorCurrentSlide = 0 Then PPTK_ColorCurrentSlide = ActiveWindow.Selection.SlideRange(1).SlideIndex
    Err.Clear
End Function

Public Function PPTK_ColorSlideTotal() As Long
    On Error Resume Next
    PPTK_ColorSlideTotal = ActivePresentation.Slides.Count
End Function

' ---------- scan ----------

Public Sub PPTK_ColorScan(ByVal fromSlide As Long, ByVal toSlide As Long, ByVal doFill As Boolean, ByVal doLine As Boolean, ByVal doFont As Boolean)
    mCount = 0
    ReDim mRgb(1 To 16)
    ReDim mFill(1 To 16)
    ReDim mLine(1 To 16)
    ReDim mFont(1 To 16)
    ReDim mSlides(1 To 16)
    mReplace = False
    PPTK_Walk fromSlide, toSlide, doFill, doLine, doFont
    PPTK_SortByUse
End Sub

Public Function PPTK_ColorCount() As Long
    PPTK_ColorCount = mCount
End Function

Public Sub PPTK_ColorInfo(ByVal index As Long, ByRef rgbValue As Long, ByRef total As Long, ByRef fills As Long, ByRef lines As Long, ByRef fonts As Long, ByRef slideList As String)
    rgbValue = mRgb(index)
    fills = mFill(index)
    lines = mLine(index)
    fonts = mFont(index)
    total = fills + lines + fonts
    slideList = PPTK_SlideListText(mSlides(index))
End Sub

' Union of the slides that contain any of the given colors ("," separated RGB values).
Public Function PPTK_ColorSlidesFor(ByVal rgbCsv As String, ByRef total As Long) As String
    Dim index As Long
    Dim merged As String
    Dim parts() As String
    Dim part As Long
    total = 0
    merged = ","
    For index = 1 To mCount
        If InStr(1, "," & rgbCsv & ",", "," & mRgb(index) & ",") > 0 Then
            total = total + mFill(index) + mLine(index) + mFont(index)
            parts = Split(mSlides(index), ",")
            For part = LBound(parts) To UBound(parts)
                If Len(parts(part)) > 0 Then
                    If InStr(1, merged, "," & parts(part) & ",") = 0 Then merged = merged & parts(part) & ","
                End If
            Next part
        End If
    Next index
    PPTK_ColorSlidesFor = PPTK_SlideListText(merged)
End Function

Public Function PPTK_ColorSkipText() As String
    Dim notes As String
    notes = ""
    If mSkipPictures > 0 Then notes = notes & ", " & mSkipPictures & " picture/media"
    If mSkipCharts > 0 Then notes = notes & ", " & mSkipCharts & " chart"
    If mSkipSmartArt > 0 Then notes = notes & ", " & mSkipSmartArt & " SmartArt"
    If mSkipObjects > 0 Then notes = notes & ", " & mSkipObjects & " OLE/ink"
    If mSkipEffects > 0 Then notes = notes & ", " & mSkipEffects & " gradient/pattern/picture fill"
    If mFailed > 0 Then notes = notes & ", " & mFailed & " unreadable"
    If Len(notes) > 0 Then PPTK_ColorSkipText = "Skipped: " & Mid$(notes, 3) Else PPTK_ColorSkipText = ""
End Function

' ---------- replace ----------

' Replaces every use of the given RGB colors with targetRgb in the scope and
' color types. Returns the number of properties changed.
' failed is the number of matching properties that could not be written.
Public Function PPTK_ColorReplace(ByVal fromSlide As Long, ByVal toSlide As Long, ByVal doFill As Boolean, ByVal doLine As Boolean, ByVal doFont As Boolean, ByVal rgbCsv As String, ByVal targetRgb As Long, ByRef slideList As String, ByRef failed As Long) As Long
    mReplace = True
    mSources = "," & rgbCsv & ","
    mTarget = targetRgb
    mChanged = 0
    mWriteFailed = 0
    mChangedSlides = ","
    On Error Resume Next
    Application.StartNewUndoEntry
    On Error GoTo 0
    PPTK_Walk fromSlide, toSlide, doFill, doLine, doFont
    mReplace = False
    slideList = PPTK_SlideListText(mChangedSlides)
    failed = mWriteFailed
    PPTK_ColorReplace = mChanged
End Function

' ---------- walker (shared by scan and replace) ----------

Private Sub PPTK_Walk(ByVal fromSlide As Long, ByVal toSlide As Long, ByVal doFill As Boolean, ByVal doLine As Boolean, ByVal doFont As Boolean)
    Dim index As Long
    Dim shapeIndex As Long
    Dim sld As Slide
    mDoFill = doFill
    mDoLine = doLine
    mDoFont = doFont
    mFailed = 0
    mSkipPictures = 0
    mSkipCharts = 0
    mSkipSmartArt = 0
    mSkipObjects = 0
    mSkipEffects = 0
    If fromSlide < 1 Then fromSlide = 1
    If toSlide > ActivePresentation.Slides.Count Then toSlide = ActivePresentation.Slides.Count
    For index = fromSlide To toSlide
        On Error GoTo SlideFailed
        Set sld = ActivePresentation.Slides(index)
        mSlideIndex = sld.SlideIndex
        For shapeIndex = 1 To sld.Shapes.Count
            PPTK_WalkShape sld.Shapes(shapeIndex)
        Next shapeIndex
NextSlide:
    Next index
    Exit Sub
SlideFailed:
    mFailed = mFailed + 1
    Resume NextSlide
End Sub

Private Sub PPTK_WalkShape(ByVal shp As Shape)
    Dim kind As Long
    Dim index As Long
    On Error GoTo Unreadable
    kind = shp.Type
    If kind = msoGroup Then
        For index = 1 To shp.GroupItems.Count
            PPTK_WalkShape shp.GroupItems(index)
        Next index
        Exit Sub
    End If
    Select Case kind
        Case msoPicture, msoLinkedPicture, msoMedia
            mSkipPictures = mSkipPictures + 1
            Exit Sub
        Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoInk, msoInkComment
            mSkipObjects = mSkipObjects + 1
            Exit Sub
    End Select
    If shp.HasChart Then
        mSkipCharts = mSkipCharts + 1
        Exit Sub
    End If
    If shp.HasSmartArt Then
        mSkipSmartArt = mSkipSmartArt + 1
        Exit Sub
    End If
    If shp.HasTable Then
        PPTK_WalkTable shp.Table
        Exit Sub
    End If
    If PPTK_HoldsPicture(shp) Then
        mSkipPictures = mSkipPictures + 1
        Exit Sub
    End If
    ' Lines and connectors draw no fill, whatever Fill.Visible reports.
    If mDoFill And kind <> msoLine And Not shp.Connector Then PPTK_VisitFill shp.Fill
    If mDoLine Then PPTK_VisitLine shp.Line
    If mDoFont Then PPTK_VisitText shp
    Exit Sub
Unreadable:
    mFailed = mFailed + 1
End Sub

Private Function PPTK_HoldsPicture(ByVal shp As Shape) As Boolean
    Dim contained As Long
    On Error GoTo Clean
    PPTK_HoldsPicture = False
    If shp.Type <> msoPlaceholder Then Exit Function
    contained = shp.PlaceholderFormat.ContainedType
    PPTK_HoldsPicture = (contained = msoPicture Or contained = msoLinkedPicture Or contained = msoMedia _
        Or contained = msoEmbeddedOLEObject Or contained = msoLinkedOLEObject)
Clean:
End Function

' A merged area is visited once: its cells share one top-left corner.
Private Sub PPTK_WalkTable(ByVal tbl As Table)
    Dim r As Long
    Dim c As Long
    Dim seen As String
    Dim key As String
    On Error GoTo Unreadable
    seen = "|"
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            key = Format$(tbl.Cell(r, c).Shape.Left, "0.0") & ";" & Format$(tbl.Cell(r, c).Shape.Top, "0.0")
            If InStr(1, seen, "|" & key & "|") = 0 Then
                seen = seen & key & "|"
                If mDoFill Then PPTK_VisitFill tbl.Cell(r, c).Shape.Fill
                If mDoFont Then PPTK_VisitText tbl.Cell(r, c).Shape
            End If
        Next c
    Next r
    Exit Sub
Unreadable:
    mFailed = mFailed + 1
End Sub

Private Sub PPTK_VisitFill(ByVal fmt As FillFormat)
    On Error GoTo Unreadable
    If fmt.Visible <> msoTrue Then Exit Sub
    If fmt.Type <> msoFillSolid Then
        mSkipEffects = mSkipEffects + 1
        Exit Sub
    End If
    PPTK_VisitColor fmt.ForeColor, KindFill, True
    Exit Sub
Unreadable:
    mFailed = mFailed + 1
End Sub

Private Sub PPTK_VisitLine(ByVal fmt As LineFormat)
    On Error GoTo Unreadable
    If fmt.Visible <> msoTrue Then Exit Sub
    PPTK_VisitColor fmt.ForeColor, KindLine, True
    Exit Sub
Unreadable:
    mFailed = mFailed + 1
End Sub

' Runs are read first and then visited through their own character ranges:
' writing a color can merge a run into its neighbour, which renumbers Runs.
Private Sub PPTK_VisitText(ByVal shp As Shape)
    Dim rng As TextRange
    Dim index As Long
    Dim runCount As Long
    Dim starts() As Long
    Dim lengths() As Long
    Dim runColor As ColorFormat
    Dim lastRgb As Long
    Dim haveLast As Boolean
    Dim rgbValue As Long
    On Error GoTo Unreadable
    If Not shp.HasTextFrame Then Exit Sub
    If Not shp.TextFrame.HasText Then Exit Sub
    Set rng = shp.TextFrame.TextRange
    runCount = rng.Runs.Count
    If runCount < 1 Then Exit Sub
    ReDim starts(1 To runCount)
    ReDim lengths(1 To runCount)
    For index = 1 To runCount
        starts(index) = rng.Runs(index).Start
        lengths(index) = rng.Runs(index).Length
    Next index
    haveLast = False
    For index = 1 To runCount
        Set runColor = rng.Characters(starts(index), lengths(index)).Font.Color
        If PPTK_ColorUsable(runColor) Then
            rgbValue = runColor.RGB
            ' Adjacent runs with the same color are one use; replace still writes each run.
            PPTK_VisitColor runColor, KindFont, Not (haveLast And rgbValue = lastRgb)
            lastRgb = rgbValue
            haveLast = True
        Else
            haveLast = False
        End If
    Next index
    Exit Sub
Unreadable:
    mFailed = mFailed + 1
End Sub

' A color is usable when PowerPoint reports its displayed RGB. Theme colors,
' Lighter/Darker steps and style tints/shades all report the displayed value
' (verified against rendered pixels), so they are found and matched like RGB.
Private Function PPTK_ColorUsable(ByVal fmt As ColorFormat) As Boolean
    Dim probe As Long
    On Error GoTo Unreadable
    PPTK_ColorUsable = False
    probe = fmt.RGB
    PPTK_ColorUsable = True
    Exit Function
Unreadable:
    mFailed = mFailed + 1
End Function

Private Sub PPTK_VisitColor(ByVal fmt As ColorFormat, ByVal kind As Long, ByVal countIt As Boolean)
    Dim rgbValue As Long
    On Error GoTo Unreadable
    If kind <> KindFont Then
        If Not PPTK_ColorUsable(fmt) Then Exit Sub
    End If
    rgbValue = fmt.RGB
    If mReplace Then
        If InStr(1, mSources, "," & rgbValue & ",") > 0 Then
            On Error GoTo WriteFailed
            fmt.RGB = mTarget
            On Error GoTo Unreadable
            If countIt Then mChanged = mChanged + 1
            If InStr(1, mChangedSlides, "," & mSlideIndex & ",") = 0 Then mChangedSlides = mChangedSlides & mSlideIndex & ","
        End If
    ElseIf countIt Then
        PPTK_Record rgbValue, kind
    End If
    Exit Sub
WriteFailed:
    mWriteFailed = mWriteFailed + 1
    Exit Sub
Unreadable:
    mFailed = mFailed + 1
End Sub

Private Sub PPTK_Record(ByVal rgbValue As Long, ByVal kind As Long)
    Dim index As Long
    For index = 1 To mCount
        If mRgb(index) = rgbValue Then Exit For
    Next index
    If index > mCount Then
        mCount = mCount + 1
        If mCount > UBound(mRgb) Then
            ReDim Preserve mRgb(1 To mCount * 2)
            ReDim Preserve mFill(1 To mCount * 2)
            ReDim Preserve mLine(1 To mCount * 2)
            ReDim Preserve mFont(1 To mCount * 2)
            ReDim Preserve mSlides(1 To mCount * 2)
        End If
        index = mCount
        mRgb(index) = rgbValue
        mFill(index) = 0
        mLine(index) = 0
        mFont(index) = 0
        mSlides(index) = ","
    End If
    Select Case kind
        Case KindFill: mFill(index) = mFill(index) + 1
        Case KindLine: mLine(index) = mLine(index) + 1
        Case KindFont: mFont(index) = mFont(index) + 1
    End Select
    If InStr(1, mSlides(index), "," & mSlideIndex & ",") = 0 Then mSlides(index) = mSlides(index) & mSlideIndex & ","
End Sub

' Most used first; ties keep discovery order.
Private Sub PPTK_SortByUse()
    Dim i As Long
    Dim j As Long
    For i = 2 To mCount
        j = i
        Do While j > 1
            If mFill(j) + mLine(j) + mFont(j) > mFill(j - 1) + mLine(j - 1) + mFont(j - 1) Then
                PPTK_SwapEntry j, j - 1
                j = j - 1
            Else
                Exit Do
            End If
        Loop
    Next i
End Sub

Private Sub PPTK_SwapEntry(ByVal a As Long, ByVal b As Long)
    Dim n As Long
    Dim t As String
    n = mRgb(a): mRgb(a) = mRgb(b): mRgb(b) = n
    n = mFill(a): mFill(a) = mFill(b): mFill(b) = n
    n = mLine(a): mLine(a) = mLine(b): mLine(b) = n
    n = mFont(a): mFont(a) = mFont(b): mFont(b) = n
    t = mSlides(a): mSlides(a) = mSlides(b): mSlides(b) = t
End Sub

' ",3,1,7," -> "1, 3, 7"
Private Function PPTK_SlideListText(ByVal setText As String) As String
    Dim parts() As String
    Dim numbers() As Long
    Dim count As Long
    Dim i As Long
    Dim j As Long
    Dim t As Long
    Dim result As String
    parts = Split(setText, ",")
    ReDim numbers(0 To UBound(parts) + 1)
    count = 0
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            numbers(count) = CLng(parts(i))
            count = count + 1
        End If
    Next i
    For i = 1 To count - 1
        j = i
        Do While j > 0
            If numbers(j) < numbers(j - 1) Then
                t = numbers(j): numbers(j) = numbers(j - 1): numbers(j - 1) = t
                j = j - 1
            Else
                Exit Do
            End If
        Loop
    Next i
    result = ""
    For i = 0 To count - 1
        If i > 0 Then result = result & ", "
        result = result & numbers(i)
    Next i
    PPTK_SlideListText = result
End Function

' ---------- color text helpers ----------

Public Function PPTK_ColorHex(ByVal rgbValue As Long) As String
    PPTK_ColorHex = "#" & Right$("0" & Hex$(rgbValue And &HFF&), 2) & _
        Right$("0" & Hex$((rgbValue \ &H100&) And &HFF&), 2) & _
        Right$("0" & Hex$((rgbValue \ &H10000) And &HFF&), 2)
End Function

' "#795DE5", "795de5" or "#79F" -> RGB value; False when the text is not a color.
Public Function PPTK_ParseHex(ByVal text As String, ByRef rgbValue As Long) As Boolean
    Dim hexText As String
    Dim i As Long
    PPTK_ParseHex = False
    hexText = UCase$(Trim$(text))
    If Left$(hexText, 1) = "#" Then hexText = Mid$(hexText, 2)
    If Len(hexText) = 3 Then hexText = Mid$(hexText, 1, 1) & Mid$(hexText, 1, 1) & Mid$(hexText, 2, 1) & Mid$(hexText, 2, 1) & Mid$(hexText, 3, 1) & Mid$(hexText, 3, 1)
    If Len(hexText) <> 6 Then Exit Function
    For i = 1 To 6
        If InStr(1, "0123456789ABCDEF", Mid$(hexText, i, 1)) = 0 Then Exit Function
    Next i
    rgbValue = RGB(CLng("&H" & Mid$(hexText, 1, 2)), CLng("&H" & Mid$(hexText, 3, 2)), CLng("&H" & Mid$(hexText, 5, 2)))
    PPTK_ParseHex = True
End Function

' 0..1, for choosing readable text on a swatch.
Public Function PPTK_ColorLuminance(ByVal rgbValue As Long) As Double
    PPTK_ColorLuminance = (0.299 * (rgbValue And &HFF&) + 0.587 * ((rgbValue \ &H100&) And &HFF&) + 0.114 * ((rgbValue \ &H10000) And &HFF&)) / 255#
End Function

' Quick picks for New Color: PPTK cyan, PPTK ink, white, black, gray, and the
' ten Office standard colors. Fixed values: the slide master and theme are not read.
Public Function PPTK_PaletteCount() As Long
    PPTK_PaletteCount = 15
End Function

Public Function PPTK_PaletteColor(ByVal index As Long) As Long
    Select Case index
        Case 1: PPTK_PaletteColor = RGB(35, 186, 236)    ' PPTK cyan #23BAEC
        Case 2: PPTK_PaletteColor = RGB(48, 52, 59)      ' PPTK ink #30343B
        Case 3: PPTK_PaletteColor = RGB(255, 255, 255)
        Case 4: PPTK_PaletteColor = RGB(0, 0, 0)
        Case 5: PPTK_PaletteColor = RGB(127, 127, 127)
        Case 6: PPTK_PaletteColor = RGB(192, 0, 0)
        Case 7: PPTK_PaletteColor = RGB(255, 0, 0)
        Case 8: PPTK_PaletteColor = RGB(255, 192, 0)
        Case 9: PPTK_PaletteColor = RGB(255, 255, 0)
        Case 10: PPTK_PaletteColor = RGB(146, 208, 80)
        Case 11: PPTK_PaletteColor = RGB(0, 176, 80)
        Case 12: PPTK_PaletteColor = RGB(0, 176, 240)
        Case 13: PPTK_PaletteColor = RGB(0, 112, 192)
        Case 14: PPTK_PaletteColor = RGB(0, 32, 96)
        Case Else: PPTK_PaletteColor = RGB(112, 48, 160)
    End Select
End Function
