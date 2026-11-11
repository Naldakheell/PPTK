Attribute VB_Name = "ModulePPTKShortcuts"
Option Explicit

' PPTK shortcut dispatch. The Select Case list is the only external entry.
' Direct inserts: T R U O C P G M N.
' Canvas sequences call the existing align, size, distribute, and spacing procedures.
' A A copies corner roundness from the first selected rounded rectangle onto the other rounded rectangles.
' X T copies the first AutoShape type onto the other autoshapes, text boxes, and placeholders.
' Swap position reuses ObjectsSwapPosition. Text swap, merge, margins, and
' formatting are local helpers so the text module is not imported.
' Ctrl+Shift+T pastes clipboard text into the destination and keeps its formatting.
' Ctrl+Shift+] and Ctrl+Shift+[ move the selection one z-order step.
' I opens the Windows Insert Icon command. IconsPane is not used.
' E encloses a selected SVG icon in one circle, then toggles a fill on that circle.
' \ saves the selected slides as a new PPTX. Email Slides is a ribbon button, not a shortcut.
' Slide Library is not a shortcut.
' Q harmonizes the current slide onto its theme. A P, A X, and A Z arrange shapes.
' Ctrl+2/3 and Alt+2/3 add or delete a grid edge and keep the grid's outer bounds.
' The legacy Expand/Shrink edge commands and Flip use CMD tokens with no key.
' Ctrl+4 through Ctrl+7 insert a local TK placeholder.
' Ctrl+K and Ctrl+Space open the Smart Bar.

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
Private PPTK_OutlookApp As Object
Private PPTK_OutlookMail As Object
Private PPTK_OutlookStarted As Boolean

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
        Case "E"
            PPTK_EncloseIcons
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
        Case "CTRL /"
            PPTK_SwapText
        Case ". T"
            PPTK_MergeText
        Case "CTRL SHIFT C"
            PPTK_CopyFormatting
        Case "CTRL SHIFT V"
            PPTK_PasteFormatting
        Case "CTRL SHIFT T"
            PPTK_PasteContentOnly
        Case "ALT ."
            PPTK_ZeroMargins
        Case "CTRL ]"
            PPTK_ExecuteMso "FontSizeIncrease"
        Case "CTRL ["
            PPTK_ExecuteMso "FontSizeDecrease"
        Case "CTRL SHIFT ]"
            PPTK_ExecuteMso "ObjectBringForward"
        Case "CTRL SHIFT ["
            PPTK_ExecuteMso "ObjectSendBackward"
        Case "\"
            PPTK_SaveSelectedSlides
        Case "Q"
            PPTK_FixSlideColors
        Case "A P"
            PPTK_AlignProcess
        Case "A X"
            PPTK_AlignMatrix
        Case "A Z"
            PPTK_AlignCentral
        Case "CTRL 2"
            PPTK_AddRightColumn
        Case "CTRL 3"
            PPTK_AddBottomRow
        Case "ALT 2"
            PPTK_DeleteRightColumn
        Case "ALT 3"
            PPTK_DeleteBottomRow
        Case "CTRL 4"
            PPTK_AddRightTK
        Case "CTRL 5"
            PPTK_AddBottomTK
        Case "CTRL 6"
            PPTK_AddTopTK
        Case "CTRL 7"
            PPTK_AddLeftTK
        Case "CTRL K", "CTRL SPACE"
            PPTK_ShowSmartBar
        ' Internal command tokens. Smart Bar and the Ribbon use them.
        ' PPTKKeys never sends them, so these commands have no keyboard shortcut.
        Case "CMD ADD COLUMN EXPAND"
            PPTK_AddRightColumnExpand
        Case "CMD ADD ROW EXPAND"
            PPTK_AddBottomRowExpand
        Case "CMD DELETE COLUMN SHRINK"
            PPTK_DeleteRightColumnShrink
        Case "CMD DELETE ROW SHRINK"
            PPTK_DeleteBottomRowShrink
        Case "CMD FLIP SLIDE"
            PPTK_FlipCurrentSlide
        Case "CMD REPLACE COLORS"
            PPTK_ShowReplaceColors
        Case "CMD FLIP PRESENTATION"
            PPTK_FlipEntirePresentation
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

Private Sub PPTK_InsertIcon()
    Dim sld As Slide
    If Not PPTK_ActiveSlide(sld) Then Exit Sub
    PPTK_ExecuteMso "IconInsertFromFile"
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
    Dim data As Object
    On Error GoTo Clean
    text = ""
    Set data = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
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
    Dim merged As Collection
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
    ' Only shapes whose text was merged are deleted. Pictures, tables, and
    ' shapes without text stay on the slide.
    Set merged = New Collection

    For index = 2 To targets.Count
        Set source = targets(index)
        If source.HasTextFrame Then
            If source.TextFrame2.HasText Then
                Set copied = source.TextFrame2.TextRange
                If copied.Length > 0 Then
                    If destination.Length > 0 Then destination.InsertAfter vbCrLf
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
Private Sub PPTK_StartUndo()
    On Error Resume Next
    Application.StartNewUndoEntry
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

Private Sub PPTK_SaveSelectedSlides()
    Dim sourceWindow As DocumentWindow
    Dim source As Presentation
    Dim selected As SlideRange
    Dim indexes() As Long
    Dim dialog As FileDialog
    Dim suggested As String
    Dim destination As String
    Dim tempPath As String
    Dim copyPres As Presentation
    Dim previousAlerts As PpAlertLevel
    Dim capturedSaved As Boolean
    Dim haveSaved As Boolean
    Dim alertsChanged As Boolean
    Dim accepted As Boolean
    Dim failure As Long

    On Error GoTo Failed
    If Application.Presentations.Count < 1 Then Exit Sub
    Set sourceWindow = ActiveWindow
    If sourceWindow Is Nothing Then Exit Sub
    If sourceWindow.Selection.Type <> ppSelectionSlides Then
        MsgBox "Select the slides to save in the thumbnail pane first.", vbInformation, "PPTK"
        Exit Sub
    End If
    Set selected = sourceWindow.Selection.SlideRange
    If selected Is Nothing Then Exit Sub
    If selected.Count < 1 Then Exit Sub
    Set source = sourceWindow.Presentation
    If source Is Nothing Then Exit Sub

    PPTK_CollectSlideIndexes selected, indexes
    suggested = PPTK_SuggestedPath(source, indexes)
    Set dialog = Application.FileDialog(msoFileDialogSaveAs)
    dialog.InitialFileName = suggested
    If dialog.Show <> -1 Then Exit Sub
    accepted = True
    destination = PPTK_ForcePptx(CStr(dialog.SelectedItems(1)))
    If Len(destination) = 0 Then Exit Sub
    If PPTK_SameFile(source, destination) Then
        MsgBox "Choose a different file. The open presentation was not changed.", vbExclamation
        Exit Sub
    End If

    tempPath = PPTK_TempCopyPath()
    If StrComp(tempPath, destination, vbTextCompare) = 0 Then Err.Raise 5
    PPTK_DeleteIfExists tempPath
    capturedSaved = source.Saved
    haveSaved = True
    previousAlerts = Application.DisplayAlerts
    source.SaveCopyAs tempPath, ppSaveAsOpenXMLPresentation
    source.Saved = capturedSaved
    Set copyPres = Presentations.Open(tempPath, msoFalse, msoFalse, msoFalse)
    If copyPres Is Nothing Then Err.Raise 5
    If Not PPTK_IsTempName(copyPres.Name) Then Err.Raise 5
    If StrComp(copyPres.FullName, source.FullName, vbTextCompare) = 0 Then Err.Raise 5

    Application.DisplayAlerts = ppAlertsNone
    alertsChanged = True
    PPTK_KeepSlideIndexes copyPres, indexes
    Application.DisplayAlerts = previousAlerts
    alertsChanged = False
    ' SaveAs releases the temporary file while this presentation stays open.
    copyPres.SaveAs destination, ppSaveAsOpenXMLPresentation
    If Not PPTK_FileExists(destination) Then Err.Raise 5
    PPTK_DeleteIfExists tempPath
    copyPres.Saved = True
    copyPres.Close
    Set copyPres = Nothing
    PPTK_DeleteIfExists tempPath
    PPTK_RestoreSelection sourceWindow, source, indexes
    Exit Sub

Failed:
    failure = Err.Number
    On Error Resume Next
    If alertsChanged Then Application.DisplayAlerts = previousAlerts
    If haveSaved Then source.Saved = capturedSaved
    If Not copyPres Is Nothing Then
        If PPTK_IsTempName(copyPres.Name) Or StrComp(copyPres.FullName, destination, vbTextCompare) = 0 Then
            copyPres.Saved = True
            copyPres.Close
        End If
        Set copyPres = Nothing
    End If
    PPTK_DeleteIfExists tempPath
    If failure <> 1004 And failure <> 4198 And failure <> 18 Then
        MsgBox "Could not save the selected slides. The open presentation was not changed.", vbExclamation
    End If
End Sub

Public Sub PPTK_EmailSelectedSlides(Optional ByVal control As IRibbonControl)
    Dim sourceWindow As DocumentWindow
    Dim source As Presentation
    Dim selected As SlideRange
    Dim indexes() As Long
    Dim attachmentPath As String
    Dim tempPath As String
    Dim copyPres As Presentation
    Dim previousAlerts As PpAlertLevel
    Dim capturedSaved As Boolean
    Dim haveSaved As Boolean
    Dim alertsChanged As Boolean
    Dim failure As Long

    On Error GoTo Failed
    If Application.Presentations.Count < 1 Then Exit Sub
    Set sourceWindow = ActiveWindow
    If sourceWindow Is Nothing Then Exit Sub
    If sourceWindow.Selection.Type <> ppSelectionSlides Then
        MsgBox "Select one or more slides, then use Email Slides.", vbExclamation
        Exit Sub
    End If
    Set selected = sourceWindow.Selection.SlideRange
    If selected Is Nothing Then Exit Sub
    If selected.Count < 1 Then
        MsgBox "Select one or more slides, then use Email Slides.", vbExclamation
        Exit Sub
    End If
    Set source = sourceWindow.Presentation
    If source Is Nothing Then Exit Sub

    PPTK_CollectSlideIndexes selected, indexes
    attachmentPath = PPTK_EmailAttachmentPath(source, indexes)
    If Len(attachmentPath) = 0 Then Err.Raise 5
    If PPTK_SameFile(source, attachmentPath) Then Err.Raise 5

    tempPath = PPTK_TempCopyPath()
    If StrComp(tempPath, attachmentPath, vbTextCompare) = 0 Then Err.Raise 5
    PPTK_DeleteIfExists tempPath
    capturedSaved = source.Saved
    haveSaved = True
    previousAlerts = Application.DisplayAlerts
    source.SaveCopyAs tempPath, ppSaveAsOpenXMLPresentation
    source.Saved = capturedSaved
    Set copyPres = Presentations.Open(tempPath, msoFalse, msoFalse, msoFalse)
    If copyPres Is Nothing Then Err.Raise 5
    If Not PPTK_IsTempName(copyPres.Name) Then Err.Raise 5
    If StrComp(copyPres.FullName, source.FullName, vbTextCompare) = 0 Then Err.Raise 5

    Application.DisplayAlerts = ppAlertsNone
    alertsChanged = True
    PPTK_KeepSlideIndexes copyPres, indexes
    copyPres.SaveAs attachmentPath, ppSaveAsOpenXMLPresentation
    Application.DisplayAlerts = previousAlerts
    alertsChanged = False
    If Not PPTK_PathExists(attachmentPath) Then Err.Raise 5
    If copyPres.Slides.Count <> (UBound(indexes) - LBound(indexes) + 1) Then Err.Raise 5
    PPTK_DeleteIfExists tempPath
    copyPres.Saved = True
    copyPres.Close
    Set copyPres = Nothing
    PPTK_DeleteIfExists tempPath
    PPTK_RestoreSelection sourceWindow, source, indexes

    If PPTK_DisplayOutlookDraft(PPTK_EmailSubject(source), attachmentPath) Then
        ' A draft opened by starting Classic Outlook still has the temp file open.
        ' Leave that file in place. An already-running Outlook has copied it.
        If Not PPTK_OutlookStarted Then PPTK_DeleteEmailTemp attachmentPath, source
        Exit Sub
    End If
    PPTK_RevealEmailFile attachmentPath
    Exit Sub

Failed:
    failure = Err.Number
    On Error Resume Next
    If alertsChanged Then Application.DisplayAlerts = previousAlerts
    If haveSaved Then source.Saved = capturedSaved
    If Not copyPres Is Nothing Then
        If PPTK_IsTempName(copyPres.Name) Or StrComp(copyPres.FullName, attachmentPath, vbTextCompare) = 0 Then
            copyPres.Saved = True
            copyPres.Close
        End If
        Set copyPres = Nothing
    End If
    PPTK_DeleteIfExists tempPath
    If Len(attachmentPath) > 0 Then
        If PPTK_PathExists(attachmentPath) And Not PPTK_SameFile(source, attachmentPath) Then
            PPTK_RevealEmailFile attachmentPath
            Exit Sub
        End If
    End If
    If failure <> 1004 And failure <> 4198 And failure <> 18 Then
        MsgBox "Could not export the selected slides. The open presentation was not changed.", vbExclamation
    End If
End Sub

Private Sub PPTK_CollectSlideIndexes(ByVal selected As SlideRange, ByRef indexes() As Long)
    Dim index As Long
    ReDim indexes(1 To selected.Count)
    For index = 1 To selected.Count
        indexes(index) = selected(index).SlideIndex
    Next index
    PPTK_SortUnique indexes
End Sub

Private Sub PPTK_SortUnique(ByRef values() As Long)
    Dim leftIndex As Long
    Dim rightIndex As Long
    Dim hold As Long
    Dim writeAt As Long
    For leftIndex = LBound(values) To UBound(values) - 1
        For rightIndex = leftIndex + 1 To UBound(values)
            If values(rightIndex) < values(leftIndex) Then
                hold = values(leftIndex)
                values(leftIndex) = values(rightIndex)
                values(rightIndex) = hold
            End If
        Next rightIndex
    Next leftIndex
    writeAt = LBound(values)
    For rightIndex = LBound(values) + 1 To UBound(values)
        If values(rightIndex) <> values(writeAt) Then
            writeAt = writeAt + 1
            values(writeAt) = values(rightIndex)
        End If
    Next rightIndex
    If writeAt < UBound(values) Then ReDim Preserve values(LBound(values) To writeAt)
End Sub

Private Sub PPTK_KeepSlideIndexes(ByVal copyPres As Presentation, ByRef indexes() As Long)
    Dim keep() As Boolean
    Dim index As Long
    Dim kept As Long
    If Not PPTK_IsTempName(copyPres.Name) Then Err.Raise 5
    ReDim keep(1 To copyPres.Slides.Count)
    For index = LBound(indexes) To UBound(indexes)
        If indexes(index) < 1 Or indexes(index) > copyPres.Slides.Count Then Err.Raise 5
        If Not keep(indexes(index)) Then
            keep(indexes(index)) = True
            kept = kept + 1
        End If
    Next index
    If kept < 1 Then Err.Raise 5
    For index = copyPres.Slides.Count To 1 Step -1
        If Not keep(index) Then copyPres.Slides(index).Delete
    Next index
    If copyPres.Slides.Count <> kept Then Err.Raise 5
End Sub

Private Sub PPTK_RestoreSelection(ByVal sourceWindow As DocumentWindow, ByVal source As Presentation, ByRef indexes() As Long)
    Dim pick() As Variant
    Dim index As Long
    On Error Resume Next
    sourceWindow.Activate
    ReDim pick(1 To UBound(indexes) - LBound(indexes) + 1)
    For index = LBound(indexes) To UBound(indexes)
        pick(index - LBound(indexes) + 1) = indexes(index)
    Next index
    source.Slides.Range(pick).Select
End Sub

Private Function PPTK_SuggestedPath(ByVal source As Presentation, ByRef indexes() As Long) As String
    Dim baseName As String
    Dim listed As String
    Dim leaf As String
    Dim folder As String
    Dim index As Long
    baseName = PPTK_SanitizeFileName(PPTK_BaseName(source.Name))
    If Len(baseName) = 0 Then baseName = "Slides"
    listed = ""
    For index = LBound(indexes) To UBound(indexes)
        If Len(listed) > 0 Then listed = listed & ","
        listed = listed & CStr(indexes(index))
    Next index
    leaf = baseName & " (slide " & listed & ")"
    If Len(leaf) > 160 Then
        leaf = baseName & " (" & CStr(UBound(indexes) - LBound(indexes) + 1) & " slides)"
    End If
    leaf = PPTK_SanitizeFileName(leaf) & ".pptx"
    If Len(source.Path) > 0 Then
        folder = source.Path
    Else
        folder = Environ$("TEMP")
        If Len(folder) = 0 Then folder = CurDir$
    End If
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    PPTK_SuggestedPath = folder & leaf
End Function

Private Function PPTK_ForcePptx(ByVal path As String) As String
    Dim slash As Long
    Dim dot As Long
    Dim folder As String
    Dim leaf As String
    path = Trim$(path)
    If Len(path) = 0 Then Exit Function
    slash = InStrRev(path, "\")
    If slash > 0 Then
        folder = Left$(path, slash)
        leaf = Mid$(path, slash + 1)
    Else
        leaf = path
    End If
    dot = InStrRev(leaf, ".")
    If dot > 1 Then
        leaf = Left$(leaf, dot - 1)
    ElseIf dot = 1 Then
        leaf = "Slides"
    End If
    leaf = PPTK_SanitizeFileName(leaf)
    If Len(leaf) = 0 Then leaf = "Slides"
    PPTK_ForcePptx = folder & leaf & ".pptx"
End Function

Private Function PPTK_TempCopyPath() As String
    Dim folder As String
    folder = Environ$("TEMP")
    If Len(folder) = 0 Then folder = CurDir$
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    Randomize
    PPTK_TempCopyPath = folder & "PPTK-slides-" & Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd() * 1000000)) & ".pptx"
End Function

Private Function PPTK_SameFile(ByVal source As Presentation, ByVal destination As String) As Boolean
    PPTK_SameFile = False
    If Len(source.Path) = 0 Then Exit Function
    PPTK_SameFile = (StrComp(source.FullName, destination, vbTextCompare) = 0)
End Function

Private Function PPTK_IsTempName(ByVal fileName As String) As Boolean
    PPTK_IsTempName = (StrComp(Left$(fileName, 12), "PPTK-slides-", vbTextCompare) = 0)
End Function

Private Function PPTK_BaseName(ByVal fileName As String) As String
    Dim dot As Long
    dot = InStrRev(fileName, ".")
    If dot > 1 Then
        PPTK_BaseName = Left$(fileName, dot - 1)
    Else
        PPTK_BaseName = fileName
    End If
End Function

Private Function PPTK_SanitizeFileName(ByVal value As String) As String
    Dim invalidChars As String
    Dim index As Long
    invalidChars = "\/:*?""<>|"
    For index = 1 To Len(invalidChars)
        value = Replace$(value, Mid$(invalidChars, index, 1), " ")
    Next index
    PPTK_SanitizeFileName = Trim$(value)
End Function

Private Function PPTK_FileExists(ByVal path As String) As Boolean
    Dim handle As Integer
    On Error Resume Next
    handle = FreeFile
    Open path For Binary Access Read As #handle
    If Err.Number = 0 Then
        Close #handle
        PPTK_FileExists = True
    End If
End Function

Private Sub PPTK_DeleteIfExists(ByVal path As String)
    Dim slash As Long
    Dim leaf As String
    Dim attempt As Long
    slash = InStrRev(path, "\")
    If slash > 0 Then
        leaf = Mid$(path, slash + 1)
    Else
        leaf = path
    End If
    If Not PPTK_IsTempName(leaf) Then Exit Sub
    For attempt = 1 To 40
        On Error Resume Next
        If Len(Dir$(path)) = 0 Then Exit Sub
        Kill path
        If Err.Number = 0 Then Exit Sub
        Err.Clear
        DoEvents
    Next attempt
End Sub

Private Function PPTK_PathExists(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    PPTK_PathExists = (Len(Dir$(path)) > 0)
End Function

Private Function PPTK_EmailSubject(ByVal source As Presentation) As String
    Dim subjectText As String
    subjectText = Trim$(PPTK_BaseName(source.Name))
    If Len(subjectText) = 0 Then subjectText = "Slides"
    If Len(subjectText) > 200 Then subjectText = Left$(subjectText, 200)
    PPTK_EmailSubject = subjectText
End Function

Private Function PPTK_LeafName(ByVal path As String) As String
    Dim slash As Long
    slash = InStrRev(path, "\")
    If slash > 0 Then
        PPTK_LeafName = Mid$(path, slash + 1)
    Else
        PPTK_LeafName = path
    End If
End Function

Private Function PPTK_EmailAttachmentPath(ByVal source As Presentation, ByRef indexes() As Long) As String
    Dim suggested As String
    Dim leaf As String
    Dim stem As String
    Dim folder As String
    Dim candidate As String
    Dim suffix As Long
    Dim dot As Long
    suggested = PPTK_SuggestedPath(source, indexes)
    leaf = PPTK_LeafName(suggested)
    dot = InStrRev(leaf, ".")
    If dot > 1 Then
        stem = Left$(leaf, dot - 1)
    Else
        stem = "Slides"
        leaf = stem & ".pptx"
    End If
    folder = Environ$("TEMP")
    If Len(folder) = 0 Then folder = CurDir$
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    candidate = folder & leaf
    suffix = 2
    Do While PPTK_PathExists(candidate) Or PPTK_SameFile(source, candidate)
        candidate = folder & stem & " " & CStr(suffix) & ".pptx"
        suffix = suffix + 1
        If suffix > 100 Then Err.Raise 5
    Loop
    PPTK_EmailAttachmentPath = candidate
End Function

Private Function PPTK_DisplayOutlookDraft(ByVal subjectText As String, ByVal attachmentPath As String) As Boolean
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim added As Object
    Dim leaf As String

    On Error GoTo Failed
    PPTK_DisplayOutlookDraft = False
    If Not PPTK_PathExists(attachmentPath) Then Exit Function
    leaf = PPTK_LeafName(attachmentPath)

    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    Err.Clear
    On Error GoTo Failed
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
        PPTK_OutlookStarted = True
    Else
        PPTK_OutlookStarted = False
    End If
    If outlookApp Is Nothing Then Exit Function
    ' Opening the MAPI session keeps a Classic Outlook instance alive after Display.
    outlookApp.GetNamespace("MAPI")

    Set mailItem = outlookApp.CreateItem(0)
    If mailItem Is Nothing Then Exit Function
    mailItem.Subject = subjectText
    Set added = mailItem.Attachments.Add(attachmentPath)
    If added Is Nothing Then GoTo Failed
    If added.Size < 1 Then GoTo Failed
    If mailItem.Attachments.Count < 1 Then GoTo Failed
    If StrComp(added.FileName, leaf, vbTextCompare) <> 0 Then GoTo Failed
    Set PPTK_OutlookApp = outlookApp
    Set PPTK_OutlookMail = mailItem
    mailItem.Display
    PPTK_DisplayOutlookDraft = True
    Exit Function
Failed:
    On Error Resume Next
    If Not mailItem Is Nothing Then mailItem.Close 1
    PPTK_DisplayOutlookDraft = False
End Function

Private Sub PPTK_DeleteEmailTemp(ByVal path As String, ByVal source As Presentation)
    Dim folder As String
    Dim attempt As Long
    If Len(path) = 0 Then Exit Sub
    If Not source Is Nothing Then
        If PPTK_SameFile(source, path) Then Exit Sub
    End If
    folder = Environ$("TEMP")
    If Len(folder) = 0 Then Exit Sub
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    If StrComp(Left$(path, Len(folder)), folder, vbTextCompare) <> 0 Then Exit Sub
    If LCase$(Right$(path, 5)) <> ".pptx" Then Exit Sub
    For attempt = 1 To 20
        On Error Resume Next
        If Len(Dir$(path)) = 0 Then Exit Sub
        Kill path
        If Err.Number = 0 Then Exit Sub
        Err.Clear
        DoEvents
    Next attempt
End Sub

Private Function PPTK_EnsureFolder(ByVal folder As String) As Boolean
    On Error GoTo Failed
    If Len(folder) = 0 Then Exit Function
    If Len(Dir$(folder, vbDirectory)) > 0 Then
        PPTK_EnsureFolder = True
        Exit Function
    End If
    MkDir folder
    PPTK_EnsureFolder = (Len(Dir$(folder, vbDirectory)) > 0)
    Exit Function
Failed:
    PPTK_EnsureFolder = False
End Function

Private Function PPTK_EmailFallbackFolder() As String
    Dim profile As String
    Dim candidate As String
    profile = Environ$("USERPROFILE")
    If Len(profile) = 0 Then Exit Function
    candidate = profile & "\Documents\PPTK Email"
    If PPTK_EnsureFolder(candidate) Then
        PPTK_EmailFallbackFolder = candidate & "\"
        Exit Function
    End If
    candidate = profile & "\Desktop\PPTK Email"
    If PPTK_EnsureFolder(candidate) Then
        PPTK_EmailFallbackFolder = candidate & "\"
    End If
End Function

Private Sub PPTK_RevealEmailFile(ByVal attachmentPath As String)
    Dim folder As String
    Dim destination As String
    Dim leaf As String
    Dim stem As String
    Dim dot As Long
    Dim suffix As Long
    Dim copied As Boolean

    If Len(attachmentPath) = 0 Then Exit Sub
    If Not PPTK_PathExists(attachmentPath) Then Exit Sub
    folder = PPTK_EmailFallbackFolder()
    leaf = PPTK_LeafName(attachmentPath)
    destination = attachmentPath
    If Len(folder) > 0 Then
        dot = InStrRev(leaf, ".")
        If dot > 1 Then
            stem = Left$(leaf, dot - 1)
        Else
            stem = leaf
        End If
        destination = folder & leaf
        suffix = 2
        Do While PPTK_PathExists(destination)
            destination = folder & stem & " " & CStr(suffix) & ".pptx"
            suffix = suffix + 1
            If suffix > 100 Then Exit Do
        Loop
        On Error Resume Next
        FileCopy attachmentPath, destination
        copied = (Err.Number = 0 And PPTK_PathExists(destination))
        Err.Clear
        On Error GoTo 0
        If copied Then
            PPTK_DeleteEmailTemp attachmentPath, Nothing
        Else
            destination = attachmentPath
        End If
    End If
    On Error Resume Next
    Shell "explorer.exe /select,""" & destination & """", vbNormalFocus
    On Error GoTo 0
    MsgBox "Classic Outlook could not open a draft." & vbCrLf & vbCrLf & _
        "The slides were saved here:" & vbCrLf & destination & vbCrLf & vbCrLf & _
        "Attach that file to your email yourself.", vbExclamation
End Sub
