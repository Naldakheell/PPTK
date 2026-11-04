Attribute VB_Name = "ModulePPTKShortcuts"
Option Explicit

' PPTK shortcut dispatch. The Select Case list is the only external entry.
' Direct inserts: T R U O C P G M N.
' Canvas sequences call the existing align, size, distribute, and spacing procedures.
' Swap position reuses ObjectsSwapPosition. Text swap, merge, margins, and
' formatting are local helpers so the text module is not imported.
' Later tokens, including order, library, and icons, are ignored.

Private Const PPTK_TAG As String = "PPTK_SHORTCUT"
Private Const PPTK_W_STANDARD As Single = 150
Private Const PPTK_H_STANDARD As Single = 100
Private Const PPTK_W_PROCESS As Single = 150
Private Const PPTK_H_PROCESS As Single = 50
Private Const PPTK_W_TEXT As Single = 300
Private Const PPTK_H_TEXT As Single = 50
Private Const PPTK_NUDGE As Single = 12
Private Const PPTK_NUDGE_LIMIT As Long = 40

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
        Case "CTRL /"
            PPTK_SwapText
        Case ". T"
            PPTK_MergeText
        Case "CTRL SHIFT C"
            PPTK_CopyFormatting
        Case "CTRL SHIFT V"
            PPTK_PasteFormatting
        Case "ALT ."
            PPTK_ZeroMargins
        Case "CTRL ]"
            PPTK_ExecuteMso "FontSizeIncrease"
        Case "CTRL ["
            PPTK_ExecuteMso "FontSizeDecrease"
        Case Else
    End Select
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
    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight
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

Private Sub PPTK_ExecuteMso(ByVal controlID As String)
    On Error Resume Next
    Application.CommandBars.ExecuteMso controlID
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

Private Sub PPTK_SwapText()
    Dim window As DocumentWindow
    Dim text1 As String
    Dim text2 As String
    On Error GoTo Clean
    Set window = Application.ActiveWindow
    If window Is Nothing Then Exit Sub
    If Not window.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        Exit Sub
    End If

    If window.Selection.ShapeRange.Count = 2 Then
        If window.Selection.ShapeRange(1).HasTextFrame And window.Selection.ShapeRange(2).HasTextFrame Then
            text1 = window.Selection.ShapeRange(1).TextFrame.TextRange.Text
            text2 = window.Selection.ShapeRange(2).TextFrame.TextRange.Text
            window.Selection.ShapeRange(1).TextFrame.TextRange.Text = text2
            window.Selection.ShapeRange(2).TextFrame.TextRange.Text = text1
        Else
            MsgBox "Select two shapes that (can) have text."
        End If
    ElseIf window.Selection.HasChildShapeRange Then
        If window.Selection.ChildShapeRange.Count = 2 Then
            If window.Selection.ChildShapeRange(1).HasTextFrame And window.Selection.ChildShapeRange(2).HasTextFrame Then
                text1 = window.Selection.ChildShapeRange(1).TextFrame.TextRange.Text
                text2 = window.Selection.ChildShapeRange(2).TextFrame.TextRange.Text
                window.Selection.ChildShapeRange(1).TextFrame.TextRange.Text = text2
                window.Selection.ChildShapeRange(2).TextFrame.TextRange.Text = text1
            Else
                MsgBox "Select two shapes that (can) have text."
            End If
        Else
            MsgBox "Select two shapes to swap their text."
        End If
    Else
        MsgBox "Select two shapes to swap their text."
    End If
Clean:
End Sub

Private Sub PPTK_MergeText()
    Dim window As DocumentWindow
    Dim target As Shape
    Dim source As Shape
    Dim targets As ShapeRange
    Dim destination As TextRange2
    Dim copied As TextRange2
    Dim index As Long
    On Error GoTo Clean
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
    If destination.Length > 0 Then destination.InsertAfter vbCrLf

    For index = 2 To targets.Count
        Set source = targets(index)
        If source.HasTextFrame Then
            If source.TextFrame2.HasText Then
                Set copied = source.TextFrame2.TextRange
                If copied.Length > 0 Then
                    copied.Copy
                    DoEvents
                    destination.Characters(destination.Length + 1, 0).Paste
                    DoEvents
                    destination.InsertAfter vbCrLf
                End If
            End If
        End If
    Next index

    For index = targets.Count To 2 Step -1
        targets(index).Delete
    Next index
Clean:
End Sub

Private Sub PPTK_ZeroMargins()
    Dim window As DocumentWindow
    Dim index As Long
    On Error GoTo Clean
    Set window = Application.ActiveWindow
    If window Is Nothing Then Exit Sub
    If Not (window.Selection.Type = ppSelectionShapes Or window.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
        Exit Sub
    End If
    If window.Selection.HasChildShapeRange Then
        For index = 1 To window.Selection.ChildShapeRange.Count
            PPTK_ZeroMarginsShape window.Selection.ChildShapeRange(index)
        Next index
    Else
        For index = 1 To window.Selection.ShapeRange.Count
            PPTK_ZeroMarginsShape window.Selection.ShapeRange(index)
        Next index
    End If
Clean:
End Sub

Private Sub PPTK_ZeroMarginsShape(ByVal target As Shape)
    Dim child As Shape
    On Error GoTo Clean
    If target.Type = msoGroup Then
        For Each child In target.GroupItems
            PPTK_ZeroMarginsShape child
        Next child
    ElseIf target.HasTextFrame Then
        target.TextFrame.MarginBottom = 0
        target.TextFrame.MarginLeft = 0
        target.TextFrame.MarginRight = 0
        target.TextFrame.MarginTop = 0
    End If
Clean:
End Sub
