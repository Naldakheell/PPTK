Attribute VB_Name = "ModulePPTKTakeaway"
Option Explicit

' PPTK local TK behavior. Auxi does not publish what TK means, so this is
' not auxi parity and it does not summarize the slide.
'
' Each command inserts one editable text box whose text is the placeholder
' "TK". It is placed against the bounding box of the selected shapes, on the
' requested side, with a 12 point margin. The box is 140 by 36 points unless
' the slide margin requires it to move. Placement is clamped inside an 8 point
' slide margin. Existing shapes and their text are left alone.
'
' A shape selection is required. The box uses Dark 1 text, a Light 1 fill,
' and an Accent 1 outline from the current slide's theme when those colors
' can be read.

Private Const BoxWidth As Single = 140
Private Const BoxHeight As Single = 36
Private Const Gap As Single = 12
Private Const SlideMargin As Single = 8

Public Sub PPTK_AddRightTK()
    PPTK_AddTK 1
End Sub

Public Sub PPTK_AddBottomTK()
    PPTK_AddTK 2
End Sub

Public Sub PPTK_AddTopTK()
    PPTK_AddTK 3
End Sub

Public Sub PPTK_AddLeftTK()
    PPTK_AddTK 4
End Sub

Private Sub PPTK_AddTK(ByVal direction As Long)
    Dim sld As Slide
    Dim leftPos As Single
    Dim topPos As Single
    Dim rightPos As Single
    Dim bottomPos As Single
    Dim boxLeft As Single
    Dim boxTop As Single
    Dim box As Shape
    Dim slideW As Single
    Dim slideH As Single

    On Error GoTo Clean
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    If ActiveWindow.Selection.ShapeRange.Count < 1 Then Exit Sub
    On Error Resume Next
    Set sld = ActiveWindow.View.Slide
    If sld Is Nothing Then Set sld = ActiveWindow.Selection.SlideRange(1)
    On Error GoTo Clean
    If sld Is Nothing Then Exit Sub
    PPTK_SelectionBounds leftPos, topPos, rightPos, bottomPos
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight

    Select Case direction
        Case 1
            boxLeft = rightPos + Gap
            boxTop = topPos
        Case 2
            boxLeft = leftPos
            boxTop = bottomPos + Gap
        Case 3
            boxLeft = leftPos
            boxTop = topPos - Gap - BoxHeight
        Case 4
            boxLeft = leftPos - Gap - BoxWidth
            boxTop = topPos
        Case Else
            Exit Sub
    End Select

    If boxLeft < SlideMargin Then boxLeft = SlideMargin
    If boxTop < SlideMargin Then boxTop = SlideMargin
    If boxLeft + BoxWidth > slideW - SlideMargin Then boxLeft = slideW - SlideMargin - BoxWidth
    If boxTop + BoxHeight > slideH - SlideMargin Then boxTop = slideH - SlideMargin - BoxHeight
    If boxLeft < SlideMargin Then boxLeft = SlideMargin
    If boxTop < SlideMargin Then boxTop = SlideMargin

    Set box = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, boxLeft, boxTop, BoxWidth, BoxHeight)
    box.Name = "PPTK TK"
    box.TextFrame.WordWrap = msoTrue
    box.TextFrame.AutoSize = ppAutoSizeNone
    box.TextFrame.MarginLeft = 4
    box.TextFrame.MarginRight = 4
    box.TextFrame.MarginTop = 2
    box.TextFrame.MarginBottom = 2
    box.TextFrame.TextRange.Text = "TK"
    On Error Resume Next
    box.TextFrame.TextRange.Font.Size = 14
    box.TextFrame.TextRange.Font.Color.ObjectThemeColor = msoThemeColorDark1
    box.Fill.Visible = msoTrue
    box.Fill.Solid
    box.Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
    box.Line.Visible = msoTrue
    box.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent1
    box.Line.Weight = 1
Clean:
End Sub

Private Sub PPTK_SelectionBounds(ByRef leftPos As Single, ByRef topPos As Single, ByRef rightPos As Single, ByRef bottomPos As Single)
    Dim index As Long
    Dim shp As Shape
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    leftPos = shp.Left
    topPos = shp.Top
    rightPos = shp.Left + shp.Width
    bottomPos = shp.Top + shp.Height
    For index = 2 To ActiveWindow.Selection.ShapeRange.Count
        Set shp = ActiveWindow.Selection.ShapeRange(index)
        If shp.Left < leftPos Then leftPos = shp.Left
        If shp.Top < topPos Then topPos = shp.Top
        If shp.Left + shp.Width > rightPos Then rightPos = shp.Left + shp.Width
        If shp.Top + shp.Height > bottomPos Then bottomPos = shp.Top + shp.Height
    Next index
End Sub
