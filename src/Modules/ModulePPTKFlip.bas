Attribute VB_Name = "ModulePPTKFlip"
Option Explicit

' Horizontal layout mirror of top-level slide shapes.
' newLeft = SlideWidth - Left - Width. Width, height, text, font, and
' reading direction stay as they are. Shape.Flip is used only for pictures
' and for text-free chevrons and arrow pentagons, because Flip mirrors glyphs.
'
' Groups, connectors, charts, and SmartArt are moved as one box and not flipped.
' A chevron or arrow pentagon that contains text is only moved.
' A second run restores positions. Pictures and text-free chevrons or pentagons
' flip back on the second run.

Public Sub PPTK_FlipCurrentSlide()
    Dim sld As Slide
    Dim slideWidth As Single
    Dim reason As String

    If Not PPTK_TryCurrentSlide(sld, reason) Then
        MsgBox reason, vbExclamation
        Exit Sub
    End If

    On Error GoTo Failed
    If sld.Shapes.Count = 0 Then
        MsgBox "Nothing to flip.", vbExclamation
        Exit Sub
    End If
    If Not PPTK_ReadSlideWidth(slideWidth) Then
        MsgBox "Could not read the slide size.", vbExclamation
        Exit Sub
    End If

    PPTK_StartUndo "Flip slide"
    PPTK_FlipSlide sld, slideWidth
    Exit Sub
Failed:
    MsgBox "Could not flip the current slide.", vbExclamation
End Sub

Public Sub PPTK_FlipEntirePresentation()
    Dim pres As Presentation
    Dim sld As Slide
    Dim slideWidth As Single

    On Error Resume Next
    Err.Clear
    Set pres = ActivePresentation
    If Err.Number <> 0 Or pres Is Nothing Then
        MsgBox "No presentation is open.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Failed
    If pres.Slides.Count = 0 Or PPTK_CountShapes(pres) = 0 Then
        MsgBox "Nothing to flip.", vbExclamation
        Exit Sub
    End If
    If Not PPTK_ReadSlideWidth(slideWidth) Then
        MsgBox "Could not read the slide size.", vbExclamation
        Exit Sub
    End If

    PPTK_StartUndo "Flip presentation"
    For Each sld In pres.Slides
        PPTK_FlipSlide sld, slideWidth
    Next sld
    Exit Sub
Failed:
    MsgBox "Could not flip the presentation.", vbExclamation
End Sub

Private Function PPTK_TryCurrentSlide(ByRef sld As Slide, ByRef reason As String) As Boolean
    Dim pres As Presentation
    Dim windowMissing As Boolean
    Dim viewIsNormal As Boolean

    On Error Resume Next
    PPTK_TryCurrentSlide = False
    Set sld = Nothing

    Err.Clear
    Set pres = ActivePresentation
    If Err.Number <> 0 Or pres Is Nothing Then
        reason = "No presentation is open."
        Exit Function
    End If

    Err.Clear
    windowMissing = True
    windowMissing = (ActiveWindow Is Nothing)
    If Err.Number <> 0 Or windowMissing Then
        reason = "No active window."
        Exit Function
    End If

    Err.Clear
    viewIsNormal = False
    viewIsNormal = (ActiveWindow.ViewType = ppViewNormal)
    If Err.Number <> 0 Or Not viewIsNormal Then
        reason = "Switch to Normal view."
        Exit Function
    End If

    Err.Clear
    Set sld = ActiveWindow.View.Slide
    If Err.Number <> 0 Or sld Is Nothing Then
        reason = "Could not read the current slide."
        Set sld = Nothing
        Exit Function
    End If

    PPTK_TryCurrentSlide = True
End Function

Private Function PPTK_ReadSlideWidth(ByRef slideWidth As Single) As Boolean
    On Error Resume Next
    Err.Clear
    slideWidth = ActivePresentation.PageSetup.SlideWidth
    PPTK_ReadSlideWidth = (Err.Number = 0 And slideWidth > 0)
End Function

Private Function PPTK_CountShapes(ByVal pres As Presentation) As Long
    Dim sld As Slide
    Dim total As Long
    Dim onSlide As Long

    On Error Resume Next
    total = 0
    For Each sld In pres.Slides
        Err.Clear
        onSlide = sld.Shapes.Count
        If Err.Number = 0 Then total = total + onSlide
    Next sld
    PPTK_CountShapes = total
End Function

Private Sub PPTK_StartUndo(ByVal entryName As String)
    On Error Resume Next
    Application.StartNewUndoEntry entryName
End Sub

Private Sub PPTK_FlipSlide(ByVal sld As Slide, ByVal slideWidth As Single)
    Dim index As Long
    Dim total As Long
    Dim shp As Shape

    On Error Resume Next
    total = sld.Shapes.Count
    For index = 1 To total
        Err.Clear
        Set shp = sld.Shapes(index)
        If Err.Number = 0 Then PPTK_FlipShape shp, slideWidth
    Next index
End Sub

Private Sub PPTK_FlipShape(ByVal shp As Shape, ByVal slideWidth As Single)
    Dim newLeft As Single
    Dim kind As MsoShapeType
    Dim isConnector As MsoTriState
    Dim hasText As Boolean
    Dim stopHere As Boolean

    On Error Resume Next
    If shp Is Nothing Then Exit Sub

    Err.Clear
    newLeft = slideWidth - shp.Left - shp.Width
    If Err.Number <> 0 Then Exit Sub

    Err.Clear
    shp.Left = newLeft
    If Err.Number <> 0 Then Exit Sub

    Err.Clear
    kind = shp.Type
    If Err.Number <> 0 Then Exit Sub

    ' Group: one box. Children stay in place relative to the group.
    If kind = msoGroup Then Exit Sub

    Err.Clear
    isConnector = shp.Connector
    If Err.Number <> 0 Then
        Err.Clear
    ElseIf isConnector = msoTrue Then
        Exit Sub
    End If

    If kind = msoChart Or kind = msoSmartArt Then Exit Sub

    Err.Clear
    stopHere = shp.HasChart
    If Err.Number <> 0 Then
        Err.Clear
    ElseIf stopHere Then
        Exit Sub
    End If

    Err.Clear
    stopHere = shp.HasSmartArt
    If Err.Number <> 0 Then
        Err.Clear
    ElseIf stopHere Then
        Exit Sub
    End If

    If kind = msoTable Then
        PPTK_SwapTableAlignment shp
        Exit Sub
    End If

    Err.Clear
    stopHere = shp.HasTable
    If Err.Number <> 0 Then
        Err.Clear
        stopHere = False
    End If
    If stopHere Then
        PPTK_SwapTableAlignment shp
        Exit Sub
    End If

    hasText = PPTK_HasText(shp)

    If kind = msoAutoShape Then PPTK_MirrorDirection shp, hasText

    If Not hasText Then
        If kind = msoPicture Or kind = msoLinkedPicture Then
            Err.Clear
            shp.Flip msoFlipHorizontal
        End If
    End If

    If hasText Then PPTK_SwapShapeAlignment shp
End Sub

Private Sub PPTK_MirrorDirection(ByVal shp As Shape, ByVal hasText As Boolean)
    Dim autoType As MsoAutoShapeType
    Dim mapped As MsoAutoShapeType

    On Error Resume Next
    Err.Clear
    autoType = shp.AutoShapeType
    If Err.Number <> 0 Then Exit Sub

    Select Case autoType
        Case msoShapeRightArrow
            mapped = msoShapeLeftArrow
        Case msoShapeLeftArrow
            mapped = msoShapeRightArrow
        Case msoShapeRightArrowCallout
            mapped = msoShapeLeftArrowCallout
        Case msoShapeLeftArrowCallout
            mapped = msoShapeRightArrowCallout
        Case msoShapeLeftRightArrow
            Exit Sub
        Case msoShapeChevron, msoShapePentagon
            If hasText Then Exit Sub
            Err.Clear
            shp.Flip msoFlipHorizontal
            Exit Sub
        Case Else
            Exit Sub
    End Select

    Err.Clear
    shp.AutoShapeType = mapped
End Sub

Private Function PPTK_HasText(ByVal shp As Shape) As Boolean
    Dim hasFrame As MsoTriState
    Dim hasBody As MsoTriState

    On Error Resume Next
    PPTK_HasText = False

    Err.Clear
    hasFrame = shp.HasTextFrame
    If Err.Number <> 0 Then
        ' Unreadable frame: do not Flip, so glyphs cannot be mirrored by mistake.
        PPTK_HasText = True
        Exit Function
    End If
    If hasFrame <> msoTrue Then Exit Function

    Err.Clear
    hasBody = shp.TextFrame2.HasText
    If Err.Number <> 0 Then
        PPTK_HasText = True
        Exit Function
    End If
    PPTK_HasText = (hasBody = msoTrue)
End Function

Private Sub PPTK_SwapTableAlignment(ByVal shp As Shape)
    Dim rowCount As Long
    Dim colCount As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim cellShape As Shape

    On Error Resume Next
    If PPTK_HasText(shp) Then PPTK_SwapShapeAlignment shp

    Err.Clear
    rowCount = shp.Table.Rows.Count
    colCount = shp.Table.Columns.Count
    If Err.Number <> 0 Then Exit Sub

    For rowIndex = 1 To rowCount
        For colIndex = 1 To colCount
            Err.Clear
            Set cellShape = Nothing
            Set cellShape = shp.Table.Cell(rowIndex, colIndex).Shape
            If Err.Number = 0 Then
                If Not cellShape Is Nothing Then PPTK_SwapShapeAlignment cellShape
            End If
        Next colIndex
    Next rowIndex
End Sub

Private Sub PPTK_SwapShapeAlignment(ByVal shp As Shape)
    On Error Resume Next
    If shp.HasTextFrame <> msoTrue Then Exit Sub
    PPTK_SwapRangeAlignment shp.TextFrame2.TextRange
End Sub

Private Sub PPTK_SwapRangeAlignment(ByVal textRange As TextRange2)
    Dim paraCount As Long
    Dim index As Long
    Dim para As TextRange2

    On Error Resume Next
    If textRange Is Nothing Then Exit Sub

    Err.Clear
    paraCount = textRange.Paragraphs.Count
    If Err.Number <> 0 Or paraCount < 1 Then
        Err.Clear
        PPTK_SwapOneAlignment textRange
        Exit Sub
    End If

    For index = 1 To paraCount
        Err.Clear
        Set para = textRange.Paragraphs(index)
        If Err.Number = 0 Then PPTK_SwapOneAlignment para
    Next index
End Sub

Private Sub PPTK_SwapOneAlignment(ByVal textRange As TextRange2)
    Dim align As Long

    On Error Resume Next
    If textRange Is Nothing Then Exit Sub

    Err.Clear
    align = textRange.ParagraphFormat.Alignment
    If Err.Number <> 0 Then Exit Sub
    If align = ppAlignLeft Then
        textRange.ParagraphFormat.Alignment = ppAlignRight
    ElseIf align = ppAlignRight Then
        textRange.ParagraphFormat.Alignment = ppAlignLeft
    End If
End Sub
