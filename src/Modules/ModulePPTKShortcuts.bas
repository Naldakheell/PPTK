Attribute VB_Name = "ModulePPTKShortcuts"
Option Explicit

' PPTK shortcut dispatch. The Select Case list is the only external entry.
' Unknown tokens are ignored.
' A A copies corner roundness from the first selected rounded rectangle onto the other rounded rectangles.
' X T copies the first AutoShape type onto the other autoshapes, text boxes, and placeholders.
' E encloses a selected SVG icon in one circle, then toggles a fill on that circle.
' Option+. removes text margins and fits the shape (Windows Alt+.).
' Control+2/3 and Option+2/3 add or delete a grid edge and keep the grid's outer bounds.
' Control+4 through Control+7 insert a local TK placeholder.
' The Expand/Shrink edge commands use CMD tokens with no key.
' Q fixes slide colors. A P, A X and A Z arrange shapes (Smart Align). Windows token spellings
' (CTRL /, CTRL SHIFT C/V, CTRL ]/[ and CTRL SHIFT ]/[) are accepted for the ribbon.

Private Const PPTK_TAG As String = "PPTK_SHORTCUT"
Private Const PPTK_W_STANDARD As Single = 150
Private Const PPTK_H_STANDARD As Single = 100
Private Const PPTK_W_PROCESS As Single = 150
Private Const PPTK_H_PROCESS As Single = 50
Private Const PPTK_W_TEXT As Single = 300
Private Const PPTK_H_TEXT As Single = 50
Private Const PPTK_NUDGE As Single = 12
Private Const PPTK_NUDGE_LIMIT As Long = 40
Private Const PPTK_CORNER_MAX As Single = 0.5

Public Sub PPTK_DispatchShortcut(ByVal token As String)
    Dim normalized As String
    normalized = PPTK_NormalizeToken(token)
    Select Case normalized
        Case "T"
            PPTK_InsertTextbox
        Case "R"
            PPTK_InsertRectangle
        Case "U"
            PPTK_InsertRoundedRectangle
        Case "O"
            PPTK_InsertOval
        Case "C"
            PPTK_InsertChevron
        Case "P"
            PPTK_InsertPentagon
        Case "G"
            PPTK_InsertTriangle
        Case "M"
            PPTK_InsertParallelogram
        Case "N"
            PPTK_InsertNewSlide
        Case "I"
            PPTK_InsertIcon
        Case "A L"
            PPTK_AlignEdge "L"
        Case "A R"
            PPTK_AlignEdge "R"
        Case "A T"
            PPTK_AlignEdge "T"
        Case "A B"
            PPTK_AlignEdge "B"
        Case "A M"
            PPTK_AlignEdge "M"
        Case "A C"
            PPTK_AlignEdge "C"
        Case "A A"
            PPTK_MatchCornerRoundness
        Case "X T"
            PPTK_UnifyShapeType
        Case "E"
            PPTK_EncloseIcons
        Case "CTRL 2"
            PPTK_AddRightColumn
        Case "CTRL 3"
            PPTK_AddBottomRow
        Case "OPT 2", "ALT 2"
            PPTK_DeleteRightColumn
        Case "OPT 3", "ALT 3"
            PPTK_DeleteBottomRow
        Case "CTRL 4"
            PPTK_AddRightTK
        Case "CTRL 5"
            PPTK_AddBottomTK
        Case "CTRL 6"
            PPTK_AddTopTK
        Case "CTRL 7"
            PPTK_AddLeftTK
        ' Internal command tokens with no keyboard shortcut (ribbon use).
        Case "CMD ADD COLUMN EXPAND"
            PPTK_AddRightColumnExpand
        Case "CMD ADD ROW EXPAND"
            PPTK_AddBottomRowExpand
        Case "CMD DELETE COLUMN SHRINK"
            PPTK_DeleteRightColumnShrink
        Case "CMD DELETE ROW SHRINK"
            PPTK_DeleteBottomRowShrink
        Case "S W"
            ObjectsSameWidth
        Case "S H"
            ObjectsSameHeight
        Case "S S"
            ObjectsSameHeightAndWidth
        Case "D H"
            ObjectsDistributeHorizontally
        Case "D V"
            ObjectsDistributeVertically
        Case "H +"
            ObjectsIncreaseSpacingHorizontal
        Case "H -"
            ObjectsDecreaseSpacingHorizontal
        Case "V +"
            ObjectsIncreaseSpacingVertical
        Case "V -"
            ObjectsDecreaseSpacingVertical
        Case "/"
            ObjectsSwapPosition
        Case "CMD /", "CTRL /"
            ObjectsSwapTextNoFormatting
        Case ". T"
            PPTK_MergeText
        Case "CMD SHIFT C", "CTRL SHIFT C"
            PPTK_CopyFormatting
        Case "CMD SHIFT V", "CTRL SHIFT V"
            PPTK_PasteFormatting
        Case "OPT .", "ALT ."
            PPTK_ZeroMargins
        Case "FONT +", "CTRL ]"
            PPTK_ExecuteMso "FontSizeIncrease"
        Case "FONT -", "CTRL ["
            PPTK_ExecuteMso "FontSizeDecrease"
        Case "ORDER +", "CTRL SHIFT ]"
            PPTK_ExecuteMso "ObjectBringForward"
        Case "ORDER -", "CTRL SHIFT ["
            PPTK_ExecuteMso "ObjectSendBackward"
        Case "\"
            PPTK_SaveSelectedSlidesIfSlides
        Case ";"
            ShowSlideLibrary
        Case "EMAIL SLIDES"
            EmailSelectedSlides
        Case "Q"
            PPTK_FixSlideColors
        Case "A P"
            PPTK_AlignProcess
        Case "A X"
            PPTK_AlignMatrix
        Case "A Z"
            PPTK_AlignCentral
        ' Internal command tokens (ribbon, Smart Bar). The keyboard helper never sends them.
        Case "CMD SELECT FILL"
            PPTK_SelectSimilarFill
        Case "CMD SELECT OUTLINE"
            PPTK_SelectSimilarOutline
        Case "CMD SELECT FONT"
            PPTK_SelectSimilarFont
        Case "CMD SELECT FONT COLOR"
            PPTK_SelectSimilarFontColor
        Case "CMD SELECT TYPE"
            PPTK_SelectSimilarType
        Case "CMD SELECT SIZE"
            PPTK_SelectSimilarSize
        Case "CMD SELECT FONT SIZE"
            PPTK_SelectSimilarFontSize
        Case "CMD SELECT WIDTH"
            PPTK_SelectSimilarWidth
        Case "CMD SELECT HEIGHT"
            PPTK_SelectSimilarHeight
        Case "CMD ALIGN REFERENCE"
            PPTK_ToggleAlignReference
        Case "CMD FLIP SLIDE"
            PPTK_FlipCurrentSlide
        Case "CMD REPLACE COLORS"
            PPTK_ShowReplaceColors
        Case "CMD FLIP PRESENTATION"
            PPTK_FlipEntirePresentation
        Case "CMD SHIFT T", "CTRL SHIFT T"
            PPTK_PasteContentOnly
        Case "CTRL K", "CTRL SPACE"
            PPTK_ShowSmartBar
        Case "CMD REPLACE FONTS"
            PPTK_ShowReplaceFonts
        Case Else
    End Select
End Sub

Private Sub PPTK_SaveSelectedSlidesIfSlides()
    Dim slidesSelected As Boolean
    slidesSelected = False
    On Error Resume Next
    slidesSelected = (ActiveWindow.Selection.Type = ppSelectionSlides)
    On Error GoTo 0
    If slidesSelected Then SaveSelectedSlides
End Sub

Private Sub PPTK_ExecuteMso(ByVal controlID As String)
    On Error Resume Next
    Dim bars As Object
    Set bars = Application.CommandBars
    bars.ExecuteMso controlID
End Sub

Private Sub PPTK_CopyFormatting()
    On Error GoTo Clean
    If ActiveWindow Is Nothing Then Exit Sub
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    If ActiveWindow.Selection.HasChildShapeRange Then
        ActiveWindow.Selection.ChildShapeRange.PickUp
    Else
        ActiveWindow.Selection.ShapeRange.PickUp
    End If
Clean:
End Sub

Private Sub PPTK_PasteFormatting()
    On Error GoTo Clean
    If ActiveWindow Is Nothing Then Exit Sub
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    If ActiveWindow.Selection.HasChildShapeRange Then
        ActiveWindow.Selection.ChildShapeRange.Apply
    Else
        ActiveWindow.Selection.ShapeRange.Apply
    End If
Clean:
End Sub

Private Function PPTK_NormalizeToken(ByVal token As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String
    Dim pendingSpace As Boolean
    out = ""
    pendingSpace = False
    For i = 1 To Len(token)
        ch = UCase$(Mid$(token, i, 1))
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            If Len(out) > 0 Then pendingSpace = True
        Else
            If pendingSpace Then
                out = out & " "
                pendingSpace = False
            End If
            out = out & ch
        End If
    Next i
    PPTK_NormalizeToken = out
End Function

Private Function PPTK_ActiveSlide(ByRef sld As Slide) As Boolean
    Dim picked As Slide
    PPTK_ActiveSlide = False
    Set sld = Nothing
    On Error GoTo Clean
    If Application.Presentations.Count < 1 Then Exit Function
    If ActiveWindow Is Nothing Then Exit Function
    If ActiveWindow.ViewType <> ppViewNormal Then Exit Function

    On Error Resume Next
    Set picked = ActiveWindow.View.Slide
    If picked Is Nothing Then
        If ActiveWindow.Selection.Type = ppSelectionSlides Then
            If ActiveWindow.Selection.SlideRange.Count >= 1 Then
                Set picked = ActiveWindow.Selection.SlideRange(1)
            End If
        End If
    End If
    On Error GoTo Clean
    If picked Is Nothing Then Exit Function
    Set sld = picked
    PPTK_ActiveSlide = True
Clean:
End Function

Private Function PPTK_TokenCovers(ByVal sld As Slide, ByVal token As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthPts As Single, ByVal heightPts As Single) As Boolean
    Dim i As Long
    Dim shp As Shape
    PPTK_TokenCovers = False
    On Error GoTo Clean
    For i = 1 To sld.Shapes.Count
        Set shp = sld.Shapes(i)
        If StrComp(shp.Tags(PPTK_TAG), token, vbTextCompare) = 0 Then
            If Abs(shp.Left - leftPos) <= 0.5 Then
                If Abs(shp.Top - topPos) <= 0.5 Then
                    If Abs(shp.Width - widthPts) <= 0.5 Then
                        If Abs(shp.Height - heightPts) <= 0.5 Then
                            PPTK_TokenCovers = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
Clean:
End Function

Private Sub PPTK_PlaceOrigin(ByVal sld As Slide, ByVal token As String, ByVal widthPts As Single, ByVal heightPts As Single, ByRef leftPos As Single, ByRef topPos As Single)
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim nudgeIndex As Long
    Dim candidateLeft As Single
    Dim candidateTop As Single
    slideWidth = ActivePresentation.PageSetup.SlideWidth
    slideHeight = ActivePresentation.PageSetup.SlideHeight
    leftPos = (slideWidth - widthPts) / 2
    topPos = (slideHeight - heightPts) / 2
    If leftPos < 0 Then leftPos = 0
    If topPos < 0 Then topPos = 0
    For nudgeIndex = 0 To PPTK_NUDGE_LIMIT
        candidateLeft = leftPos + (PPTK_NUDGE * nudgeIndex)
        candidateTop = topPos + (PPTK_NUDGE * nudgeIndex)
        If candidateLeft > slideWidth Or candidateTop > slideHeight Then Exit For
        If Not PPTK_TokenCovers(sld, token, candidateLeft, candidateTop, widthPts, heightPts) Then
            leftPos = candidateLeft
            topPos = candidateTop
            Exit Sub
        End If
    Next nudgeIndex
    leftPos = candidateLeft
    topPos = candidateTop
End Sub

Private Sub PPTK_InsertShape(ByVal shapeType As MsoAutoShapeType, ByVal widthPts As Single, ByVal heightPts As Single, ByVal token As String)
    Dim sld As Slide
    Dim shp As Shape
    Dim leftPos As Single
    Dim topPos As Single
    If Not PPTK_ActiveSlide(sld) Then Exit Sub
    On Error GoTo Clean
    PPTK_PlaceOrigin sld, token, widthPts, heightPts, leftPos, topPos
    Set shp = sld.Shapes.AddShape(shapeType, leftPos, topPos, widthPts, heightPts)
    shp.Tags.Add PPTK_TAG, token
    shp.Select
Clean:
End Sub

Private Sub PPTK_TryEnterTextEdit(ByVal shp As Shape)
    On Error GoTo Clean
    shp.Select
    If Not shp.HasTextFrame Then Exit Sub
    shp.TextFrame.TextRange.Select
    If ActiveWindow.Selection.Type = ppSelectionText Then Exit Sub
    On Error Resume Next
    shp.TextFrame2.TextRange.Select
Clean:
End Sub

Private Sub PPTK_InsertTextbox()
    Dim sld As Slide
    Dim shp As Shape
    Dim leftPos As Single
    Dim topPos As Single
    If Not PPTK_ActiveSlide(sld) Then Exit Sub
    On Error GoTo Clean
    PPTK_PlaceOrigin sld, "T", PPTK_W_TEXT, PPTK_H_TEXT, leftPos, topPos
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, PPTK_W_TEXT, PPTK_H_TEXT)
    shp.Tags.Add PPTK_TAG, "T"
    PPTK_TryEnterTextEdit shp
Clean:
End Sub

Private Sub PPTK_InsertRectangle()
    PPTK_InsertShape msoShapeRectangle, PPTK_W_STANDARD, PPTK_H_STANDARD, "R"
End Sub

Private Sub PPTK_InsertRoundedRectangle()
    PPTK_InsertShape msoShapeRoundedRectangle, PPTK_W_STANDARD, PPTK_H_STANDARD, "U"
End Sub

Private Sub PPTK_InsertOval()
    PPTK_InsertShape msoShapeOval, PPTK_W_STANDARD, PPTK_H_STANDARD, "O"
End Sub

Private Sub PPTK_InsertChevron()
    PPTK_InsertShape msoShapeChevron, PPTK_W_PROCESS, PPTK_H_PROCESS, "C"
End Sub

Private Sub PPTK_InsertPentagon()
    PPTK_InsertShape msoShapePentagon, PPTK_W_PROCESS, PPTK_H_PROCESS, "P"
End Sub

Private Sub PPTK_InsertTriangle()
    PPTK_InsertShape msoShapeIsoscelesTriangle, PPTK_W_STANDARD, PPTK_H_STANDARD, "G"
End Sub

Private Sub PPTK_InsertParallelogram()
    PPTK_InsertShape msoShapeParallelogram, PPTK_W_STANDARD, PPTK_H_STANDARD, "M"
End Sub

Private Sub PPTK_InsertIcon()
    Dim sld As Slide
    If Not PPTK_ActiveSlide(sld) Then Exit Sub
    PPTK_ExecuteMso "IconsPane"
End Sub

Private Sub PPTK_InsertNewSlide()
    Dim sld As Slide
    Dim layout As CustomLayout
    Dim created As Slide
    If Not PPTK_ActiveSlide(sld) Then Exit Sub
    On Error GoTo Clean
    Set layout = sld.CustomLayout
    Set created = ActivePresentation.Slides.AddSlide(sld.SlideIndex + 1, layout)
    created.Select
Clean:
End Sub

Private Function PPTK_SelectedShapes(ByRef selected As ShapeRange) As Boolean
    Dim window As DocumentWindow
    Dim hasChildren As Boolean
    Dim selectionType As PpSelectionType
    Dim shapeCount As Long
    PPTK_SelectedShapes = False
    Set selected = Nothing
    On Error Resume Next
    Set window = ActiveWindow
    If Err.Number <> 0 Or window Is Nothing Then
        Err.Clear
        Exit Function
    End If
    selectionType = window.Selection.Type
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    If selectionType <> ppSelectionShapes Then
        Err.Clear
        Exit Function
    End If
    hasChildren = window.Selection.HasChildShapeRange
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    If hasChildren Then
        Set selected = window.Selection.ChildShapeRange
    Else
        Set selected = window.Selection.ShapeRange
    End If
    If Err.Number <> 0 Or selected Is Nothing Then
        Err.Clear
        Set selected = Nothing
        Exit Function
    End If
    shapeCount = selected.Count
    If Err.Number <> 0 Or shapeCount < 1 Then
        Err.Clear
        Set selected = Nothing
        Exit Function
    End If
    Err.Clear
    PPTK_SelectedShapes = True
End Function

Private Function PPTK_IsRoundedRectangle(ByVal target As Shape) As Boolean
    PPTK_IsRoundedRectangle = False
    On Error GoTo Clean
    If target Is Nothing Then Exit Function
    If target.Type <> msoAutoShape Then Exit Function
    If target.AutoShapeType <> msoShapeRoundedRectangle Then Exit Function
    If target.Adjustments.Count < 1 Then Exit Function
    PPTK_IsRoundedRectangle = True
Clean:
    Err.Clear
End Function

Private Sub PPTK_ApplyMatchedCorner(ByVal target As Shape, ByVal radius As Single, ByRef limited As Boolean)
    Dim shorter As Single
    Dim targetAdj As Double
    Dim wasLimited As Boolean
    On Error GoTo Clean
    shorter = target.Width
    If target.Height < shorter Then shorter = target.Height
    If shorter <= 0 Then Exit Sub
    targetAdj = CDbl(radius) / CDbl(shorter)
    wasLimited = (targetAdj > PPTK_CORNER_MAX) Or (targetAdj < 0)
    If targetAdj > PPTK_CORNER_MAX Then targetAdj = PPTK_CORNER_MAX
    If targetAdj < 0 Then targetAdj = 0
    target.Adjustments(1) = CSng(targetAdj)
    If wasLimited Then limited = True
    Exit Sub
Clean:
    Err.Clear
End Sub

' A A. The first selected rounded rectangle is the reference. PowerPoint
' draws a rounded-rectangle corner with radius = Adjustments(1) * Min(Width,
' Height) (ECMA-376 roundRect, adj pinned to 0..0.5). This was checked by
' exporting two differently sized rectangles and measuring the corners.
' Each other rounded rectangle gets the adjustment that gives the same
' radius in points. Only Adjustments(1) changes. Other shape types, including
' ordinary rectangles, are skipped. Nothing is rotated.
' The reference is the first rounded rectangle in selection order (click,
' then Shift+click the others). A marquee selection uses stacking order.
Private Sub PPTK_MatchCornerRoundness()
    Dim selected As ShapeRange
    Dim target As Shape
    Dim index As Long
    Dim referenceIndex As Long
    Dim shorter As Single
    Dim radius As Single
    Dim eligibleTargets As Long
    Dim limited As Boolean

    On Error GoTo Clean
    If Not PPTK_SelectedShapes(selected) Then
        MsgBox "Select two or more rounded rectangles. The first one sets the corner roundness.", vbExclamation, "PPTK"
        Exit Sub
    End If

    referenceIndex = 0
    For index = 1 To selected.Count
        On Error Resume Next
        Set target = Nothing
        Set target = selected(index)
        If Err.Number = 0 Then
            If PPTK_IsRoundedRectangle(target) Then referenceIndex = index
        End If
        Err.Clear
        On Error GoTo Clean
        If referenceIndex > 0 Then Exit For
    Next index
    If referenceIndex = 0 Then
        MsgBox "Select two or more rounded rectangles. The first one sets the corner roundness.", vbExclamation, "PPTK"
        Exit Sub
    End If

    On Error Resume Next
    Set target = selected(referenceIndex)
    shorter = target.Width
    If target.Height < shorter Then shorter = target.Height
    radius = target.Adjustments(1) * shorter
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "The corner roundness of the first rounded rectangle could not be read.", vbExclamation, "PPTK"
        Exit Sub
    End If
    On Error GoTo Clean

    eligibleTargets = 0
    For index = referenceIndex + 1 To selected.Count
        On Error Resume Next
        Set target = Nothing
        Set target = selected(index)
        If Err.Number = 0 Then
            If PPTK_IsRoundedRectangle(target) Then eligibleTargets = eligibleTargets + 1
        End If
        Err.Clear
        On Error GoTo Clean
    Next index
    If eligibleTargets = 0 Then
        MsgBox "Select two or more rounded rectangles. The first one sets the corner roundness.", vbExclamation, "PPTK"
        Exit Sub
    End If

    PPTK_StartUndo
    limited = False
    For index = referenceIndex + 1 To selected.Count
        On Error Resume Next
        Set target = Nothing
        Set target = selected(index)
        If Err.Number = 0 Then
            If PPTK_IsRoundedRectangle(target) Then PPTK_ApplyMatchedCorner target, radius, limited
        End If
        Err.Clear
        On Error GoTo Clean
    Next index

    If limited Then
        MsgBox "Some shapes are too small for this corner radius. Their corners were set to the maximum roundness.", vbInformation, "PPTK"
    End If
Clean:
End Sub

' One Undo step per command.
' Late bound so the module compiles even if PowerPoint for Mac lacks StartNewUndoEntry.
Private Sub PPTK_StartUndo()
    PPTK_NewUndoEntry
End Sub

' Mac replacement for Application.StartNewUndoEntry, used by every PPTK module.
' PowerPoint for Mac does not compile the early-bound call, and the late-bound call can raise
' at run time. The caller's error handler must never see that error.
Public Sub PPTK_NewUndoEntry()
    Dim app As Object
    On Error Resume Next
    Set app = Application
    app.StartNewUndoEntry
    Err.Clear
End Sub

Private Function PPTK_IsConnectorShape(ByVal target As Shape) As Boolean
    PPTK_IsConnectorShape = False
    On Error Resume Next
    If target.Connector = msoTrue Then PPTK_IsConnectorShape = True
    Err.Clear
End Function

Private Function PPTK_IsAutoShapeSource(ByVal target As Shape) As Boolean
    PPTK_IsAutoShapeSource = False
    On Error GoTo Clean
    If target Is Nothing Then Exit Function
    If PPTK_IsConnectorShape(target) Then Exit Function
    If target.Type <> msoAutoShape Then Exit Function
    PPTK_IsAutoShapeSource = True
Clean:
    Err.Clear
End Function

Private Function PPTK_CanReceiveShapeType(ByVal target As Shape) As Boolean
    Dim shapeType As MsoShapeType
    Dim blocksType As Boolean
    PPTK_CanReceiveShapeType = False
    On Error Resume Next
    If target Is Nothing Then
        Err.Clear
        Exit Function
    End If
    If PPTK_IsConnectorShape(target) Then
        Err.Clear
        Exit Function
    End If
    shapeType = target.Type
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    Select Case shapeType
        Case msoPicture, msoLinkedPicture, msoChart, msoTable, msoGroup, msoLine
            Err.Clear
            Exit Function
        Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject
            Err.Clear
            Exit Function
        Case msoMedia, msoInk, msoInkComment
            Err.Clear
            Exit Function
    End Select
    blocksType = target.HasTable
    If Err.Number <> 0 Then
        Err.Clear
        blocksType = False
    End If
    If blocksType Then
        Err.Clear
        Exit Function
    End If
    blocksType = target.HasChart
    If Err.Number <> 0 Then
        Err.Clear
        blocksType = False
    End If
    If blocksType Then
        Err.Clear
        Exit Function
    End If
    If shapeType = msoPlaceholder Then
        ' A picture, chart, table, or media placeholder would be cropped or changed.
        Dim heldType As MsoShapeType
        blocksType = False
        heldType = target.PlaceholderFormat.ContainedType
        If Err.Number <> 0 Then
            Err.Clear
        ElseIf heldType <> msoAutoShape And heldType <> msoTextBox And heldType <> msoPlaceholder Then
            blocksType = True
        End If
        If blocksType Then Exit Function
    End If
    If shapeType = msoAutoShape Or shapeType = msoTextBox Or shapeType = msoPlaceholder Then
        PPTK_CanReceiveShapeType = True
    End If
    Err.Clear
End Function

Private Function PPTK_FirstAutoShape(ByVal selected As ShapeRange, ByRef copiedType As MsoAutoShapeType) As Long
    Dim index As Long
    Dim candidate As Shape
    Dim shapeType As MsoAutoShapeType
    PPTK_FirstAutoShape = 0
    On Error Resume Next
    For index = 1 To selected.Count
        Set candidate = Nothing
        Set candidate = selected(index)
        If Err.Number = 0 Then
            If PPTK_IsAutoShapeSource(candidate) Then
                shapeType = candidate.AutoShapeType
                If Err.Number = 0 And shapeType <> msoShapeMixed And shapeType <> msoShapeNotPrimitive Then
                    copiedType = shapeType
                    PPTK_FirstAutoShape = index
                    Err.Clear
                    Exit Function
                End If
            End If
        End If
        Err.Clear
    Next index
    Err.Clear
End Function

' True when the target already has the type, so it is left untouched
' (its own corner and other adjustments stay).
Private Function PPTK_HasAutoShapeType(ByVal target As Shape, ByVal copiedType As MsoAutoShapeType) As Boolean
    Dim current As MsoAutoShapeType
    On Error Resume Next
    PPTK_HasAutoShapeType = False
    current = target.AutoShapeType
    If Err.Number = 0 Then PPTK_HasAutoShapeType = (current = copiedType)
    Err.Clear
End Function

Private Sub PPTK_AssignAutoShapeType(ByVal target As Shape, ByVal copiedType As MsoAutoShapeType)
    Dim savedLeft As Single
    Dim savedTop As Single
    Dim savedWidth As Single
    Dim savedHeight As Single
    Dim savedRotation As Single
    Dim savedLock As MsoTriState
    Dim lockKnown As Boolean
    Dim drifted As Boolean

    On Error Resume Next
    savedLeft = target.Left
    savedTop = target.Top
    savedWidth = target.Width
    savedHeight = target.Height
    savedRotation = target.Rotation
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    lockKnown = False
    savedLock = target.LockAspectRatio
    If Err.Number = 0 Then lockKnown = True
    Err.Clear

    target.AutoShapeType = copiedType
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    drifted = False
    If Abs(target.Left - savedLeft) > 0.01 Then drifted = True
    If Abs(target.Top - savedTop) > 0.01 Then drifted = True
    If Abs(target.Width - savedWidth) > 0.01 Then drifted = True
    If Abs(target.Height - savedHeight) > 0.01 Then drifted = True
    If Abs(target.Rotation - savedRotation) > 0.01 Then drifted = True
    If Err.Number <> 0 Then
        Err.Clear
        drifted = True
    End If
    If Not drifted Then
        Err.Clear
        Exit Sub
    End If

    target.LockAspectRatio = msoFalse
    Err.Clear
    target.Left = savedLeft
    Err.Clear
    target.Top = savedTop
    Err.Clear
    target.Width = savedWidth
    Err.Clear
    target.Height = savedHeight
    Err.Clear
    target.Rotation = savedRotation
    Err.Clear
    If lockKnown Then target.LockAspectRatio = savedLock
    Err.Clear
End Sub

' X T. The first selected AutoShape is the reference. Every other eligible
' shape (AutoShape, text box, text or empty placeholder) gets its
' AutoShapeType. Each target keeps its position, size, rotation, flip,
' text, fill, outline, and stacking order, because the type is changed on
' the shape itself. Targets that already have the type are not touched, so
' their own corner roundness stays. Adjustments are never copied: PowerPoint
' gives a changed shape the new type's default adjustments. A A matches
' corner roundness separately.
Private Sub PPTK_UnifyShapeType()
    Dim selected As ShapeRange
    Dim target As Shape
    Dim index As Long
    Dim referenceIndex As Long
    Dim copiedType As MsoAutoShapeType
    Dim candidateCount As Long

    On Error GoTo Clean
    If Not PPTK_SelectedShapes(selected) Then
        MsgBox "Select two or more shapes. The first AutoShape sets the shape type.", vbExclamation, "PPTK"
        Exit Sub
    End If
    If selected.Count < 2 Then
        MsgBox "Select two or more shapes. The first AutoShape sets the shape type.", vbExclamation, "PPTK"
        Exit Sub
    End If

    referenceIndex = PPTK_FirstAutoShape(selected, copiedType)
    If referenceIndex = 0 Then
        MsgBox "Select an AutoShape first. Its shape type is applied to the other shapes.", vbExclamation, "PPTK"
        Exit Sub
    End If

    candidateCount = 0
    For index = 1 To selected.Count
        If index <> referenceIndex Then
            On Error Resume Next
            Set target = Nothing
            Set target = selected(index)
            If Err.Number = 0 Then
                If PPTK_CanReceiveShapeType(target) Then candidateCount = candidateCount + 1
            End If
            Err.Clear
            On Error GoTo Clean
        End If
    Next index
    If candidateCount < 1 Then
        MsgBox "None of the other selected objects can take a shape type.", vbExclamation, "PPTK"
        Exit Sub
    End If

    PPTK_StartUndo
    For index = 1 To selected.Count
        If index <> referenceIndex Then
            On Error Resume Next
            Set target = Nothing
            Set target = selected(index)
            If Err.Number = 0 Then
                If PPTK_CanReceiveShapeType(target) Then
                    If Not PPTK_HasAutoShapeType(target, copiedType) Then PPTK_AssignAutoShapeType target, copiedType
                End If
            End If
            Err.Clear
            On Error GoTo Clean
        End If
    Next index
Clean:
End Sub

Private Sub PPTK_EncloseIcons()
    Dim window As DocumentWindow
    Dim selected As ShapeRange
    Dim hasChildren As Boolean
    Dim allIds() As Long
    Dim iconIds() As Long
    Dim iconCount As Long
    Dim index As Long
    Dim sld As Slide
    Dim icon As Shape

    On Error GoTo Clean
    Set window = ActiveWindow
    If window Is Nothing Then Exit Sub
    If window.ViewType <> ppViewNormal Then Exit Sub
    If window.Selection.Type <> ppSelectionShapes Then Exit Sub

    On Error Resume Next
    hasChildren = window.Selection.HasChildShapeRange
    If Err.Number <> 0 Then Exit Sub
    Err.Clear
    On Error GoTo Clean
    If hasChildren Then Exit Sub

    Set selected = window.Selection.ShapeRange
    If selected Is Nothing Then Exit Sub
    If selected.Count < 1 Then Exit Sub

    ReDim allIds(1 To selected.Count)
    iconCount = 0
    For index = 1 To selected.Count
        allIds(index) = selected(index).Id
        If PPTK_IsIcon(selected(index)) Then
            If TypeName(selected(index).Parent) = "Slide" Then
                iconCount = iconCount + 1
                ReDim Preserve iconIds(1 To iconCount)
                iconIds(iconCount) = selected(index).Id
                Set sld = selected(index).Parent
            End If
        End If
    Next index
    If iconCount = 0 Then Exit Sub

    For index = 1 To iconCount
        Set icon = PPTK_ShapeById(sld, iconIds(index))
        If Not icon Is Nothing Then PPTK_EncloseOne icon
    Next index
    PPTK_ReselectIds sld, allIds
Clean:
End Sub

Private Sub PPTK_EncloseOne(ByVal icon As Shape)
    Dim sld As Slide
    Dim enclosure As Shape
    Dim created As Boolean

    On Error GoTo Clean
    If Not PPTK_IsIcon(icon) Then Exit Sub
    If TypeName(icon.Parent) <> "Slide" Then Exit Sub
    Set sld = icon.Parent
    Set enclosure = PPTK_FindEnclosure(sld, icon)
    If enclosure Is Nothing Then
        created = True
    Else
        created = False
    End If
    If created Then Set enclosure = PPTK_CreateEnclosure(sld, icon)
    If enclosure Is Nothing Then Exit Sub
    PPTK_PlaceEnclosure enclosure, icon
    If created Then
        PPTK_ApplyEnclosureStyle enclosure, "OUTLINE"
    ElseIf PPTK_TagValue(enclosure, "PPTK_ENCLOSURE_STATE") = "FILL" Then
        PPTK_ApplyEnclosureStyle enclosure, "OUTLINE"
    Else
        PPTK_ApplyEnclosureStyle enclosure, "FILL"
    End If
    PPTK_PlaceBehind icon, enclosure
Clean:
End Sub

Private Function PPTK_IsIcon(ByVal target As Shape) As Boolean
    PPTK_IsIcon = False
    On Error GoTo Clean
    If target Is Nothing Then Exit Function
    If PPTK_TagValue(target, "PPTK_ENCLOSURE_FOR") <> "" Then Exit Function
    PPTK_IsIcon = (target.Type = msoGraphic Or target.Type = msoLinkedGraphic)
Clean:
End Function

Private Function PPTK_CreateEnclosure(ByVal sld As Slide, ByVal icon As Shape) As Shape
    Dim enclosure As Shape
    On Error GoTo Clean
    Set enclosure = sld.Shapes.AddShape(msoShapeOval, icon.Left, icon.Top, icon.Width, icon.Height)
    enclosure.Tags.Add "PPTK_ENCLOSURE_FOR", CStr(icon.Id)
    On Error Resume Next
    enclosure.Name = "PPTK Enclosure"
    On Error GoTo Clean
    Set PPTK_CreateEnclosure = enclosure
Clean:
End Function

Private Sub PPTK_PlaceEnclosure(ByVal enclosure As Shape, ByVal icon As Shape)
    Dim side As Single
    Dim centerX As Single
    Dim centerY As Single
    On Error GoTo Clean
    side = PPTK_EnclosureSide(icon)
    centerX = icon.Left + (icon.Width / 2)
    centerY = icon.Top + (icon.Height / 2)
    enclosure.LockAspectRatio = msoFalse
    enclosure.Rotation = 0
    enclosure.Width = side
    enclosure.Height = side
    enclosure.Left = centerX - (side / 2)
    enclosure.Top = centerY - (side / 2)
Clean:
End Sub

Private Function PPTK_EnclosureSide(ByVal icon As Shape) As Single
    Dim radians As Double
    Dim cosine As Double
    Dim sine As Double
    Dim boxWidth As Single
    Dim boxHeight As Single
    radians = icon.Rotation * 3.14159265358979 / 180#
    cosine = Abs(Cos(radians))
    sine = Abs(Sin(radians))
    boxWidth = (icon.Width * cosine) + (icon.Height * sine)
    boxHeight = (icon.Width * sine) + (icon.Height * cosine)
    If boxWidth > boxHeight Then
        PPTK_EnclosureSide = boxWidth * 1.5
    Else
        PPTK_EnclosureSide = boxHeight * 1.5
    End If
End Function

Private Sub PPTK_ApplyEnclosureStyle(ByVal enclosure As Shape, ByVal stateName As String)
    On Error GoTo Clean
    enclosure.Tags.Add "PPTK_ENCLOSURE_STATE", stateName
    enclosure.Shadow.Visible = msoFalse
    enclosure.Line.Visible = msoTrue
    enclosure.Line.Weight = 1.25
    enclosure.Line.Transparency = 0
    enclosure.Line.ForeColor.ObjectThemeColor = msoThemeColorText1
    enclosure.Fill.Solid
    If stateName = "FILL" Then
        enclosure.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
        enclosure.Fill.ForeColor.Brightness = 0.6
        enclosure.Fill.Transparency = 0
        enclosure.Fill.Visible = msoTrue
    Else
        enclosure.Fill.ForeColor.RGB = RGB(255, 255, 255)
        enclosure.Fill.Transparency = 1
        enclosure.Fill.Visible = msoFalse
    End If
Clean:
End Sub

Private Sub PPTK_PlaceBehind(ByVal icon As Shape, ByVal enclosure As Shape)
    Dim guard As Long
    On Error GoTo Clean
    guard = 0
    Do While enclosure.ZOrderPosition > icon.ZOrderPosition
        enclosure.ZOrder msoSendBackward
        guard = guard + 1
        If guard > 200 Then Exit Do
    Loop
Clean:
End Sub

Private Function PPTK_FindEnclosure(ByVal sld As Slide, ByVal icon As Shape) As Shape
    Dim candidate As Shape
    Dim best As Shape
    Dim bestDistance As Double
    Dim centerX As Double
    Dim centerY As Double
    Dim distance As Double
    On Error GoTo Clean
    bestDistance = 1E+99
    centerX = icon.Left + (icon.Width / 2)
    centerY = icon.Top + (icon.Height / 2)
    For Each candidate In sld.Shapes
        If candidate.Id <> icon.Id Then
            If PPTK_TagValue(candidate, "PPTK_ENCLOSURE_FOR") = CStr(icon.Id) Then
                distance = ((candidate.Left + (candidate.Width / 2)) - centerX) ^ 2
                distance = distance + ((candidate.Top + (candidate.Height / 2)) - centerY) ^ 2
                If best Is Nothing Then
                    Set best = candidate
                    bestDistance = distance
                ElseIf distance < bestDistance Then
                    Set best = candidate
                    bestDistance = distance
                End If
            End If
        End If
    Next candidate
    Set PPTK_FindEnclosure = best
Clean:
End Function

Private Function PPTK_ShapeById(ByVal sld As Slide, ByVal shapeId As Long) As Shape
    Dim candidate As Shape
    On Error GoTo Clean
    For Each candidate In sld.Shapes
        If candidate.Id = shapeId Then
            Set PPTK_ShapeById = candidate
            Exit Function
        End If
    Next candidate
Clean:
End Function

Private Function PPTK_TagValue(ByVal target As Shape, ByVal name As String) As String
    Dim index As Long
    PPTK_TagValue = ""
    On Error GoTo Clean
    For index = 1 To target.Tags.Count
        If target.Tags.Name(index) = name Then
            PPTK_TagValue = target.Tags.Value(index)
            Exit Function
        End If
    Next index
Clean:
End Function

Private Sub PPTK_ReselectIds(ByVal sld As Slide, ByRef shapeIds() As Long)
    Dim index As Long
    Dim candidate As Shape
    Dim first As Boolean
    On Error GoTo Clean
    first = True
    For index = LBound(shapeIds) To UBound(shapeIds)
        Set candidate = PPTK_ShapeById(sld, shapeIds(index))
        If Not candidate Is Nothing Then
            If first Then
                candidate.Select
                first = False
            Else
                candidate.Select False
            End If
        End If
    Next index
Clean:
End Sub

' Alt+. Sets Left, Right, Top, and Bottom text margins to 0 through both
' TextFrame and TextFrame2. Works on a shape selection, on a text cursor or
' text selection (the shape being edited), on groups (every member), and
' on tables: selected cells only, or every cell when the whole table is
' selected or no cell reports a selection. Then the shape, or the table's
' affected columns and rows, shrink to the text and keep their center.
Private Sub PPTK_ZeroMargins()
    Dim window As DocumentWindow
    Dim targets As ShapeRange
    Dim selectionType As PpSelectionType
    Dim hasChildren As Boolean
    Dim index As Long

    On Error Resume Next
    Set window = Application.ActiveWindow
    If Err.Number <> 0 Or window Is Nothing Then Exit Sub
    selectionType = window.Selection.Type
    If Err.Number <> 0 Then Exit Sub
    If Not (selectionType = ppSelectionShapes Or selectionType = ppSelectionText) Then
        MsgBox "Select shapes, text, or table cells to remove their margins.", vbExclamation, "PPTK"
        Exit Sub
    End If

    hasChildren = window.Selection.HasChildShapeRange
    If Err.Number <> 0 Then
        Err.Clear
        hasChildren = False
    End If
    If hasChildren Then
        Set targets = window.Selection.ChildShapeRange
    Else
        Set targets = window.Selection.ShapeRange
    End If
    If Err.Number <> 0 Or targets Is Nothing Then
        Err.Clear
        Exit Sub
    End If

    PPTK_StartUndo
    For index = 1 To targets.Count
        Err.Clear
        PPTK_ZeroMarginsShape targets(index)
    Next index
    Err.Clear
End Sub

Private Sub PPTK_ZeroMarginsShape(ByVal target As Shape)
    Dim child As Shape
    Dim cellShape As Shape
    Dim tableObj As Table
    Dim tableCell As Cell
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim rowCount As Long
    Dim columnCount As Long
    Dim shapeType As MsoShapeType
    Dim hasTable As Boolean
    Dim hasFrame As Boolean
    Dim cellSelected As Boolean
    Dim selectedFound As Boolean
    Dim selectedBroken As Boolean
    Dim columnHit() As Boolean
    Dim rowHit() As Boolean

    On Error Resume Next
    If target Is Nothing Then Exit Sub
    shapeType = target.Type
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    If shapeType = msoGroup Then
        For Each child In target.GroupItems
            If Err.Number = 0 Then
                If Not child Is Nothing Then PPTK_ZeroMarginsShape child
            End If
            Err.Clear
        Next child
        Err.Clear
        Exit Sub
    End If

    hasTable = target.HasTable
    If Err.Number <> 0 Then
        Err.Clear
        hasTable = False
    End If
    If hasTable Then
        Set tableObj = target.Table
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
        rowCount = tableObj.Rows.Count
        columnCount = tableObj.Columns.Count
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
        ReDim columnHit(1 To columnCount)
        ReDim rowHit(1 To rowCount)

        selectedFound = False
        selectedBroken = False
        For rowIndex = 1 To rowCount
            For columnIndex = 1 To columnCount
                Set tableCell = tableObj.Cell(rowIndex, columnIndex)
                If Err.Number <> 0 Then
                    Err.Clear
                Else
                    cellSelected = tableCell.Selected
                    If Err.Number <> 0 Then
                        Err.Clear
                        selectedBroken = True
                        Exit For
                    ElseIf cellSelected Then
                        selectedFound = True
                    End If
                End If
            Next columnIndex
            If selectedBroken Then Exit For
        Next rowIndex

        For rowIndex = 1 To rowCount
            For columnIndex = 1 To columnCount
                Set tableCell = tableObj.Cell(rowIndex, columnIndex)
                If Err.Number <> 0 Then
                    Err.Clear
                Else
                    Set cellShape = Nothing
                    If selectedBroken Or Not selectedFound Then
                        Set cellShape = tableCell.Shape
                    Else
                        cellSelected = tableCell.Selected
                        If Err.Number = 0 Then
                            If cellSelected Then Set cellShape = tableCell.Shape
                        End If
                    End If
                    If Err.Number = 0 Then
                        If Not cellShape Is Nothing Then
                            PPTK_ZeroTextMargins cellShape
                            columnHit(columnIndex) = True
                            rowHit(rowIndex) = True
                        End If
                    End If
                    Err.Clear
                End If
            Next columnIndex
        Next rowIndex
        Err.Clear
        PPTK_FitTableToText target, columnHit, rowHit
        Err.Clear
        Exit Sub
    End If

    hasFrame = target.HasTextFrame
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    If hasFrame Then
        PPTK_ZeroTextMargins target
        PPTK_FitShapeToText target
    End If
    Err.Clear
End Sub

' Shapes that are fitted to their text after the margins are removed.
' Other outlines (ovals, chevrons, arrows) keep their size, because their
' text area is smaller than their box and the text would touch the outline.
Private Function PPTK_FitsToText(ByVal target As Shape) As Boolean
    Dim shapeType As MsoShapeType
    Dim outline As MsoAutoShapeType
    PPTK_FitsToText = False
    On Error Resume Next
    shapeType = target.Type
    If Err.Number <> 0 Then Exit Function
    If shapeType = msoTextBox Or shapeType = msoPlaceholder Then
        PPTK_FitsToText = True
    ElseIf shapeType = msoAutoShape Then
        outline = target.AutoShapeType
        If Err.Number = 0 Then PPTK_FitsToText = (outline = msoShapeRectangle Or outline = msoShapeRoundedRectangle)
    End If
    Err.Clear
End Function

' Shrinks the shape to the text it holds, keeping its center. The width
' becomes the widest line plus 1 pt, so no line wraps again. If the text
' gets taller (a line wrapped), the width grows back 1 pt at a time, up to
' the original width. Lines.Count is not used: it can miss wrapped lines.
Private Sub PPTK_FitShapeToText(ByVal target As Shape)
    Dim body As TextRange2
    Dim centerX As Single
    Dim centerY As Single
    Dim oldWidth As Single
    Dim newWidth As Single
    Dim textHeight As Single
    Dim savedLock As MsoTriState
    Dim pass As Long

    On Error Resume Next
    If Not PPTK_FitsToText(target) Then Exit Sub
    If target.TextFrame2.HasText <> msoTrue Then Exit Sub
    If target.TextFrame2.Orientation <> msoTextOrientationHorizontal Then Exit Sub
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    Set body = target.TextFrame2.TextRange
    centerX = target.Left + target.Width / 2
    centerY = target.Top + target.Height / 2
    oldWidth = target.Width
    textHeight = body.BoundHeight
    savedLock = target.LockAspectRatio
    target.LockAspectRatio = msoFalse
    target.TextFrame2.AutoSize = msoAutoSizeNone

    newWidth = body.BoundWidth + 1
    If newWidth < oldWidth Then
        target.Width = newWidth
        For pass = 1 To 20
            If body.BoundHeight <= textHeight + 0.5 Then Exit For
            target.Width = target.Width + 1
        Next pass
        If body.BoundHeight > textHeight + 0.5 Then target.Width = oldWidth
    End If
    target.Height = body.BoundHeight
    target.Left = centerX - target.Width / 2
    target.Top = centerY - target.Height / 2
    target.LockAspectRatio = savedLock
    Err.Clear
End Sub

' Shrinks the columns and rows that hold processed cells to their text,
' keeping the table's center. A column becomes its widest cell line plus
' 1 pt; a column with no text keeps its width. A row is set to 1 pt and
' PowerPoint raises it to the smallest height that shows its text. A table
' with merged cells keeps its size.
Private Sub PPTK_FitTableToText(ByVal target As Shape, ByRef columnHit() As Boolean, ByRef rowHit() As Boolean)
    Dim tableObj As Table
    Dim cellText As TextRange2
    Dim centerX As Single
    Dim centerY As Single
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim widest As Single
    Dim oldWidth As Single
    Dim heights() As Single
    Dim wrapped As Boolean
    Dim pass As Long

    On Error Resume Next
    Set tableObj = target.Table
    If Err.Number <> 0 Then Exit Sub
    For rowIndex = 1 To tableObj.Rows.Count
        For columnIndex = 1 To tableObj.Columns.Count
            If Abs(tableObj.Cell(rowIndex, columnIndex).Shape.Width - tableObj.Columns(columnIndex).Width) > 1 Then Exit Sub
            If Abs(tableObj.Cell(rowIndex, columnIndex).Shape.Height - tableObj.Rows(rowIndex).Height) > 1 Then Exit Sub
        Next columnIndex
    Next rowIndex
    If Err.Number <> 0 Then Exit Sub

    centerX = target.Left + target.Width / 2
    centerY = target.Top + target.Height / 2
    ReDim heights(1 To tableObj.Rows.Count)

    For columnIndex = 1 To tableObj.Columns.Count
        If columnHit(columnIndex) Then
            widest = 0
            For rowIndex = 1 To tableObj.Rows.Count
                Set cellText = tableObj.Cell(rowIndex, columnIndex).Shape.TextFrame2.TextRange
                heights(rowIndex) = cellText.BoundHeight
                If cellText.Length > 0 Then
                    If cellText.BoundWidth > widest Then widest = cellText.BoundWidth
                End If
            Next rowIndex
            oldWidth = tableObj.Columns(columnIndex).Width
            If widest > 0 And widest + 1 < oldWidth Then
                tableObj.Columns(columnIndex).Width = widest + 1
                For pass = 1 To 21
                    wrapped = False
                    For rowIndex = 1 To tableObj.Rows.Count
                        If tableObj.Cell(rowIndex, columnIndex).Shape.TextFrame2.TextRange.BoundHeight > heights(rowIndex) + 0.5 Then wrapped = True
                    Next rowIndex
                    If Not wrapped Then Exit For
                    If pass = 21 Then
                        tableObj.Columns(columnIndex).Width = oldWidth
                    Else
                        tableObj.Columns(columnIndex).Width = tableObj.Columns(columnIndex).Width + 1
                    End If
                Next pass
            End If
            Err.Clear
        End If
    Next columnIndex

    For rowIndex = 1 To tableObj.Rows.Count
        If rowHit(rowIndex) Then tableObj.Rows(rowIndex).Height = 1
        Err.Clear
    Next rowIndex

    target.Left = centerX - target.Width / 2
    target.Top = centerY - target.Height / 2
    Err.Clear
End Sub

Private Sub PPTK_ZeroTextMargins(ByVal target As Shape)
    On Error Resume Next
    If target Is Nothing Then Exit Sub
    target.TextFrame.MarginLeft = 0
    Err.Clear
    target.TextFrame.MarginRight = 0
    Err.Clear
    target.TextFrame.MarginTop = 0
    Err.Clear
    target.TextFrame.MarginBottom = 0
    Err.Clear
    target.TextFrame2.MarginLeft = 0
    Err.Clear
    target.TextFrame2.MarginRight = 0
    Err.Clear
    target.TextFrame2.MarginTop = 0
    Err.Clear
    target.TextFrame2.MarginBottom = 0
    Err.Clear
End Sub

' . T, the Merge text ribbon button and Feature Search all use this (Windows behavior).
Public Sub PPTK_MergeText()
    Dim window As DocumentWindow
    Dim target As Shape
    Dim source As Shape
    Dim targets As ShapeRange
    Dim destination As TextRange2
    Dim copied As TextRange2
    Dim index As Long
    Dim merged As Collection
    Dim separator As String
    On Error GoTo Clean
#If Mac Then
    ' PowerPoint for Mac stores vbCrLf as LF + paragraph mark, and the LF shows as an extra line.
    ' Windows treats vbCrLf as one paragraph break, so a lone CR gives the same result here.
    separator = vbCr
#Else
    separator = vbCrLf
#End If
    Set window = Application.ActiveWindow
    If window Is Nothing Then Exit Sub
    If Not window.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        Exit Sub
    End If
    If window.Selection.ShapeRange.Count = 1 Then
        MsgBox "Please select more than one shape."
        Exit Sub
    End If

    Set targets = window.Selection.ShapeRange
    Set target = targets(1)
    If Not target.HasTextFrame Then
        MsgBox "The first selected shape has no textframe."
        Exit Sub
    End If

    Set destination = target.TextFrame2.TextRange
    ' Only shapes whose text was merged are deleted. Pictures, tables, and
    ' shapes without text stay on the slide.
    Set merged = New Collection

    For index = 2 To targets.Count
        Set source = targets(index)
        If source.HasTextFrame Then
            If source.TextFrame2.HasText Then
                Set copied = source.TextFrame2.TextRange
                If copied.Length > 0 Then
                    If destination.Length > 0 Then
#If Mac Then
                        ' Mac: the range pasted last already ends with its paragraph mark; a second
                        ' separator would leave an empty paragraph between merged texts.
                        If Right$(destination.Text, 1) <> vbCr Then destination.InsertAfter separator
#Else
                        destination.InsertAfter separator
#End If
                    End If
                    copied.Copy
                    DoEvents
                    destination.Characters(destination.Length + 1, 0).Paste
                    DoEvents
                    merged.Add source
                End If
            End If
        End If
    Next index

    ' A pasted text range brings its closing paragraph mark; drop the empty last paragraph.
    If merged.Count > 0 And destination.Length > 0 Then
        If Right$(destination.Text, 1) = vbCr Then destination.Characters(destination.Length, 1).Delete
    End If

    For index = merged.Count To 1 Step -1
        merged(index).Delete
    Next index
Clean:
End Sub

' Cmd+Shift+T (Windows Ctrl+Shift+T): paste clipboard text into the selected text or shape, keeping its formatting.
Private Sub PPTK_PasteContentOnly()
    Dim window As DocumentWindow
    Dim selected As ShapeRange
    Dim text As String
    Dim hasChildren As Boolean

    On Error GoTo Clean
    Set window = ActiveWindow
    If window Is Nothing Then Exit Sub
    If Not PPTK_ClipboardText(text) Then Exit Sub

    If window.Selection.Type = ppSelectionText Then
        PPTK_ReplaceTextKeepingFormat window.Selection.TextRange, text
        Exit Sub
    End If
    If window.Selection.Type <> ppSelectionShapes Then Exit Sub

    On Error Resume Next
    hasChildren = window.Selection.HasChildShapeRange
    If Err.Number <> 0 Then Exit Sub
    Err.Clear
    If hasChildren Then
        Set selected = window.Selection.ChildShapeRange
    Else
        Set selected = window.Selection.ShapeRange
    End If
    If Err.Number <> 0 Then Exit Sub
    Err.Clear
    On Error GoTo Clean

    If selected Is Nothing Then Exit Sub
    If selected.Count <> 1 Then Exit Sub
    If Not selected(1).HasTextFrame Then Exit Sub
    PPTK_ReplaceTextKeepingFormat selected(1).TextFrame.TextRange, text
Clean:
End Sub

Private Sub PPTK_ReplaceTextKeepingFormat(ByVal target As TextRange, ByVal text As String)
    On Error GoTo Clean
    If target Is Nothing Then Exit Sub
    If target.Length = 0 Then
        target.InsertAfter text
    Else
        target.Text = text
    End If
Clean:
End Sub

Private Function PPTK_ClipboardText(ByRef text As String) As Boolean
#If Mac Then
    ' PowerPoint for Mac: the COM moniker fails ("Automation error"); MSForms.DataObject works (Unicode too).
    Dim data As MSForms.DataObject
    On Error GoTo Clean
    text = ""
    Set data = New MSForms.DataObject
#Else
    Dim data As Object
    On Error GoTo Clean
    text = ""
    Set data = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
#End If
    data.GetFromClipboard
    text = CStr(data.GetText())
    If Len(text) >= 2 Then
        If Right$(text, 2) = vbCrLf Then text = Left$(text, Len(text) - 2)
    End If
    If Len(text) >= 1 Then
        If Right$(text, 1) = vbCr Or Right$(text, 1) = vbLf Then text = Left$(text, Len(text) - 1)
    End If
    PPTK_ClipboardText = Len(text) > 0
Clean:
End Function

' PowerPoint for Mac draws UserForms at 72 dpi, Windows at 96 dpi: the same form is 25% smaller on Mac.
' Zooms the contents to 133% and grows the window by the same factor, so PPTK windows match Windows.
' Called by each PPTK form right after it sets its size. The frame (title bar) is measured at 100%.
' PowerPoint for Mac, zoomed forms: centred and right-aligned label text is laid out at the zoomed size but
' clipped at the unzoomed one, so it is drawn off-centre, cut or not at all. The text is drawn left-aligned
' instead, at the position where the requested alignment would put it inside boxLeft..boxLeft+boxWidth.
Public Sub PPTK_MacPlaceText(ByVal target As Object, ByVal boxLeft As Single, ByVal boxWidth As Single, ByVal align As Long)
#If Mac Then
    Dim keepTop As Single
    Dim keepHeight As Single
    Dim textWidth As Single
    On Error Resume Next
    keepTop = target.Top
    keepHeight = target.Height
    target.WordWrap = False
    target.TextAlign = 1
    target.AutoSize = True
    target.AutoSize = False
    textWidth = target.Width
    If textWidth > boxWidth Then textWidth = boxWidth
    If align = 2 Then
        target.Left = boxLeft + (boxWidth - textWidth) / 2
    ElseIf align = 3 Then
        target.Left = boxLeft + boxWidth - textWidth
    Else
        target.Left = boxLeft
    End If
    target.Width = textWidth + 2
    target.Top = keepTop
    target.Height = keepHeight
    Err.Clear
#End If
End Sub

Public Sub PPTK_MacScaleForm(ByVal target As Object)
#If Mac Then
    Const MacZoom As Single = 133
    Dim fullW As Single
    Dim fullH As Single
    Dim frameW As Single
    Dim frameH As Single
    On Error Resume Next
    If target.Zoom <> 100 Then target.Zoom = 100
    fullW = target.Width
    fullH = target.Height
    frameW = fullW - target.InsideWidth
    frameH = fullH - target.InsideHeight
    If frameW < 0 Or frameW > 48 Then frameW = 0
    If frameH < 0 Or frameH > 80 Then frameH = 0
    target.Zoom = MacZoom
    target.Width = frameW + (fullW - frameW) * MacZoom / 100
    target.Height = frameH + (fullH - frameH) * MacZoom / 100
    ' Labels keep the text layout of the old zoom (right-aligned captions such as "H +" vanish)
    ' until their size changes once. A 1 pt nudge and back redoes the layout.
    Dim item As Object
    Dim keepW As Single
    For Each item In target.Controls
        keepW = item.Width
        item.Width = keepW + 1
        item.Width = keepW
    Next item
    Err.Clear
#End If
End Sub

' Ribbon "Email Slides" (Windows v2 callback name). On Mac the existing Instrumenta EmailSelectedSlides makes the
' Outlook draft (through the Instrumenta AppleScript plugin); the guard and message match Windows.
Public Sub PPTK_EmailSelectedSlides(Optional ByVal control As IRibbonControl)
    On Error Resume Next
    If Application.Presentations.Count < 1 Then Exit Sub
    If ActiveWindow.Selection.Type <> ppSelectionSlides Then
        MsgBox "Select one or more slides, then use Email Slides.", vbExclamation
        Exit Sub
    End If
    If ActiveWindow.Selection.SlideRange.Count < 1 Then
        MsgBox "Select one or more slides, then use Email Slides.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    EmailSelectedSlides
End Sub
