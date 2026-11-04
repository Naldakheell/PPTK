Attribute VB_Name = "ModulePPTKShortcuts"
Option Explicit

' PPTK shortcut dispatch. The Select Case list is the only external entry.
' Direct inserts: T R U O C P G M N.
' Canvas sequences call the existing align, size, distribute, and spacing procedures.
' A A copies the first selected shape's rotation onto the other selected shapes.
' Swap position reuses ObjectsSwapPosition. Text swap, merge, margins, and
' formatting are local helpers so the text module is not imported.
' Ctrl+Shift+] and Ctrl+Shift+[ move the selection one z-order step.
' I opens the Windows Insert Icon command. IconsPane is not used.
' \ saves the selected slides as a new PPTX. Library and email are ignored.

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
            ObjectsAlignShapeAngle
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
        Case "CTRL SHIFT ]"
            PPTK_ExecuteMso "ObjectBringForward"
        Case "CTRL SHIFT ["
            PPTK_ExecuteMso "ObjectSendBackward"
        Case "\"
            PPTK_SaveSelectedSlides
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
    If sourceWindow.Selection.Type <> ppSelectionSlides Then Exit Sub
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
