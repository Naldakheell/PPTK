Attribute VB_Name = "ModulePPTKShortcuts"
Option Explicit

' PPTK shortcut dispatch. The Select Case list is the only external entry.
' Unknown tokens are ignored.
' A A copies corner roundness from the first selected rounded rectangle onto the other rounded rectangles.
' X T copies the first AutoShape type onto the other autoshapes, text boxes, and placeholders.

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
            ObjectsAlignLefts
        Case "A R"
            ObjectsAlignRights
        Case "A T"
            ObjectsAlignTops
        Case "A B"
            ObjectsAlignBottoms
        Case "A M"
            ObjectsAlignMiddles
        Case "A C"
            ObjectsAlignCenters
        Case "A A"
            PPTK_MatchCornerRoundness
        Case "X T"
            PPTK_UnifyShapeType
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
        Case "CMD /"
            ObjectsSwapTextNoFormatting
        Case ". T"
            ObjectsTextMerge
        Case "CMD SHIFT C"
            PPTK_CopyFormatting
        Case "CMD SHIFT V"
            PPTK_PasteFormatting
        Case "OPT ."
            ObjectsMarginsToZero
        Case "FONT +"
            PPTK_ExecuteMso "FontSizeIncrease"
        Case "FONT -"
            PPTK_ExecuteMso "FontSizeDecrease"
        Case "ORDER +"
            PPTK_ExecuteMso "ObjectBringForward"
        Case "ORDER -"
            PPTK_ExecuteMso "ObjectSendBackward"
        Case "\"
            PPTK_SaveSelectedSlidesIfSlides
        Case ";"
            ShowSlideLibrary
        Case "EMAIL SLIDES"
            EmailSelectedSlides
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
