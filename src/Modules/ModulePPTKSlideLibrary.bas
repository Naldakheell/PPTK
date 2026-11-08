Attribute VB_Name = "ModulePPTKSlideLibrary"
Option Explicit

' Local slide library. Each entry is one PPTX plus a JPEG of its first slide.
' The folder is remembered with PPTK settings. Save Selected Slides is not used.

Private Const PPTK_LIB_APP As String = "PPTK"
Private Const PPTK_LIB_SECTION As String = "SlideLibrary"
Private Const PPTK_LIB_KEY As String = "Folder"

Public Function PPTK_LibraryFolder() As String
    Dim stored As String
    stored = GetSetting(PPTK_LIB_APP, PPTK_LIB_SECTION, PPTK_LIB_KEY, "")
    If Len(stored) = 0 Then stored = PPTK_LibDefaultFolder()
    PPTK_LibraryFolder = PPTK_LibNormalizeFolder(stored)
End Function

Public Function PPTK_LibraryUseFolder(ByVal folder As String) As Boolean
    folder = PPTK_LibNormalizeFolder(folder)
    If Not PPTK_LibEnsureFolder(folder) Then Exit Function
    If Not PPTK_LibFolderWritable(folder) Then Exit Function
    SaveSetting PPTK_LIB_APP, PPTK_LIB_SECTION, PPTK_LIB_KEY, folder
    PPTK_LibraryUseFolder = True
End Function

Public Function PPTK_LibraryPickFolder() As Boolean
    Dim dialog As FileDialog
    Dim picked As String
    On Error GoTo Failed
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.Title = "Slide Library folder"
    dialog.AllowMultiSelect = False
    If dialog.Show <> -1 Then Exit Function
    picked = CStr(dialog.SelectedItems(1))
    PPTK_LibraryPickFolder = PPTK_LibraryUseFolder(picked)
    Exit Function
Failed:
    PPTK_LibraryPickFolder = False
End Function

Public Sub PPTK_SaveToLibrary(Optional ByVal control As IRibbonControl)
    Dim suggested As String
    Dim chosen As String
    If Not PPTK_LibHasSlideSelection() Then
        MsgBox "Select one or more slides, then use Save to Library.", vbExclamation
        Exit Sub
    End If
    suggested = PPTK_LibraryDefaultName()
    chosen = InputBox("Name this library entry.", "Save to Library", suggested)
    If StrPtr(chosen) = 0 Then Exit Sub
    PPTK_SaveToLibraryAs chosen
End Sub

Public Function PPTK_SaveToLibraryAs(ByVal entryName As String) As String
    Dim sourceWindow As DocumentWindow
    Dim source As Presentation
    Dim selected As SlideRange
    Dim indexes() As Long
    Dim folder As String
    Dim stem As String
    Dim libraryPath As String
    Dim previewPath As String
    Dim tempPath As String
    Dim copyPres As Presentation
    Dim previousAlerts As PpAlertLevel
    Dim capturedSaved As Boolean
    Dim haveSaved As Boolean
    Dim alertsChanged As Boolean
    Dim createdLibrary As Boolean
    Dim failure As Long

    On Error GoTo Failed
    If Not PPTK_LibHasSlideSelection() Then
        MsgBox "Select one or more slides, then use Save to Library.", vbExclamation
        Exit Function
    End If
    Set sourceWindow = ActiveWindow
    Set selected = sourceWindow.Selection.SlideRange
    Set source = sourceWindow.Presentation
    folder = PPTK_LibraryFolder()
    If Not PPTK_LibEnsureFolder(folder) Or Not PPTK_LibFolderWritable(folder) Then
        MsgBox "The slide library folder is not available:" & vbCrLf & folder & vbCrLf & vbCrLf & _
            "Choose a folder you can save to.", vbExclamation
        If Not PPTK_LibraryPickFolder() Then Exit Function
        folder = PPTK_LibraryFolder()
    End If

    PPTK_LibCollectIndexes selected, indexes
    stem = PPTK_LibSanitize(entryName)
    If Len(stem) = 0 Then stem = PPTK_LibraryDefaultName()
    If Len(stem) > 80 Then stem = Left$(stem, 80)
    stem = PPTK_LibSanitize(stem)
    If Len(stem) = 0 Then stem = "Slides"
    libraryPath = PPTK_LibUniquePath(folder, stem)
    If Len(libraryPath) = 0 Then Err.Raise 5
    If PPTK_LibSameFile(source, libraryPath) Then Err.Raise 5
    previewPath = PPTK_LibPreviewPath(libraryPath)

    tempPath = PPTK_LibTempPath()
    If StrComp(tempPath, libraryPath, vbTextCompare) = 0 Then Err.Raise 5
    PPTK_LibDeleteTemp tempPath
    capturedSaved = source.Saved
    haveSaved = True
    previousAlerts = Application.DisplayAlerts
    source.SaveCopyAs tempPath, ppSaveAsOpenXMLPresentation
    source.Saved = capturedSaved
    Set copyPres = Presentations.Open(tempPath, msoFalse, msoFalse, msoFalse)
    If copyPres Is Nothing Then Err.Raise 5
    If Not PPTK_LibIsTempName(copyPres.Name) Then Err.Raise 5
    If StrComp(copyPres.FullName, source.FullName, vbTextCompare) = 0 Then Err.Raise 5

    Application.DisplayAlerts = ppAlertsNone
    alertsChanged = True
    PPTK_LibKeepIndexes copyPres, indexes
    copyPres.SaveAs libraryPath, ppSaveAsOpenXMLPresentation
    createdLibrary = True
    If copyPres.Slides.Count <> (UBound(indexes) - LBound(indexes) + 1) Then Err.Raise 5
    On Error Resume Next
    copyPres.Slides(1).Export previewPath, "JPG", 640, 360
    Err.Clear
    On Error GoTo Failed
    Application.DisplayAlerts = previousAlerts
    alertsChanged = False
    copyPres.Saved = True
    copyPres.Close
    Set copyPres = Nothing
    PPTK_LibDeleteTemp tempPath
    PPTK_LibRestoreSelection sourceWindow, source, indexes
    If Len(Dir$(libraryPath)) = 0 Then Err.Raise 5
    PPTK_SaveToLibraryAs = libraryPath
    Exit Function

Failed:
    failure = Err.Number
    On Error Resume Next
    If alertsChanged Then Application.DisplayAlerts = previousAlerts
    If haveSaved Then source.Saved = capturedSaved
    If Not copyPres Is Nothing Then
        If PPTK_LibIsTempName(copyPres.Name) Or StrComp(copyPres.FullName, libraryPath, vbTextCompare) = 0 Then
            copyPres.Saved = True
            copyPres.Close
        End If
        Set copyPres = Nothing
    End If
    PPTK_LibDeleteTemp tempPath
    If createdLibrary Then
        If Len(Dir$(libraryPath)) > 0 Then Kill libraryPath
        If Len(previewPath) > 0 Then
            If Len(Dir$(previewPath)) > 0 Then Kill previewPath
        End If
    End If
    If failure <> 1004 And failure <> 4198 And failure <> 18 Then
        MsgBox "Could not save the library entry. The open presentation was not changed.", vbExclamation
    End If
End Function

Public Function PPTK_LibraryDefaultName() As String
    Dim selected As SlideRange
    Dim listed As String
    Dim index As Long
    Dim textValue As String
    If Not PPTK_LibHasSlideSelection() Then
        PPTK_LibraryDefaultName = "Slides"
        Exit Function
    End If
    Set selected = ActiveWindow.Selection.SlideRange
    textValue = PPTK_LibFirstText(selected(1))
    textValue = PPTK_LibSanitize(textValue)
    If Len(textValue) > 40 Then textValue = Trim$(Left$(textValue, 40))
    If Len(textValue) > 0 Then
        PPTK_LibraryDefaultName = textValue
        Exit Function
    End If
    listed = ""
    For index = 1 To selected.Count
        If Len(listed) > 0 Then listed = listed & ", "
        listed = listed & CStr(selected(index).SlideIndex)
    Next index
    PPTK_LibraryDefaultName = "Slides " & listed
End Function

Public Sub PPTK_OpenSlideLibrary(Optional ByVal control As IRibbonControl)
    PPTKSlideLibraryForm.Show vbModal
End Sub

Public Function PPTK_LibraryInsert(ByVal libraryPath As String, ByVal useSourceFormatting As Boolean) As Boolean
    Dim sourceWindow As DocumentWindow
    Dim source As Presentation
    Dim libraryPres As Presentation
    Dim destinationDesign As Design
    Dim previousAlerts As PpAlertLevel
    Dim alertsChanged As Boolean
    Dim insertAfter As Long
    Dim beforeCount As Long
    Dim added As Long
    Dim pick() As Variant
    Dim index As Long

    On Error GoTo Failed
    If Application.Presentations.Count < 1 Then
        MsgBox "Open a presentation before inserting library slides.", vbExclamation
        Exit Function
    End If
    If Len(Dir$(libraryPath)) = 0 Then
        MsgBox "That library entry is missing:" & vbCrLf & libraryPath, vbExclamation
        Exit Function
    End If
    Set sourceWindow = ActiveWindow
    If sourceWindow Is Nothing Then Err.Raise 5
    Set source = sourceWindow.Presentation
    If source Is Nothing Then Err.Raise 5
    If StrComp(source.FullName, libraryPath, vbTextCompare) = 0 Then
        MsgBox "The library entry is the open presentation. Choose another entry.", vbExclamation
        Exit Function
    End If
    insertAfter = PPTK_LibInsertAfter(sourceWindow, source)
    beforeCount = source.Slides.Count
    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = ppAlertsNone
    alertsChanged = True

    Set libraryPres = Presentations.Open(libraryPath, msoTrue, msoFalse, msoFalse)
    If libraryPres Is Nothing Then Err.Raise 5
    If libraryPres.Slides.Count < 1 Then Err.Raise 5
    added = libraryPres.Slides.Count
    If insertAfter >= 1 And insertAfter <= source.Slides.Count Then
        Set destinationDesign = source.Slides(insertAfter).Design
    ElseIf source.Slides.Count >= 1 Then
        Set destinationDesign = source.Slides(1).Design
    End If
    libraryPres.Saved = True
    libraryPres.Close
    Set libraryPres = Nothing
    sourceWindow.Activate
    source.Slides.InsertFromFile libraryPath, insertAfter, 1, added
    If Not useSourceFormatting Then
        If Not destinationDesign Is Nothing Then
            For index = 1 To added
                source.Slides(insertAfter + index).Design = destinationDesign
            Next index
        End If
    End If

    Application.DisplayAlerts = previousAlerts
    alertsChanged = False
    added = source.Slides.Count - beforeCount
    If added < 1 Then Err.Raise 5
    ReDim pick(1 To added)
    For index = 1 To added
        pick(index) = insertAfter + index
    Next index
    On Error Resume Next
    sourceWindow.Activate
    source.Slides.Range(pick).Select
    On Error GoTo 0
    PPTK_LibraryInsert = True
    Exit Function

Failed:
    On Error Resume Next
    If alertsChanged Then Application.DisplayAlerts = previousAlerts
    If Not libraryPres Is Nothing Then
        libraryPres.Saved = True
        libraryPres.Close
        Set libraryPres = Nothing
    End If
    If Not sourceWindow Is Nothing Then sourceWindow.Activate
    MsgBox "Could not insert that library entry. The library file was not changed.", vbExclamation
    PPTK_LibraryInsert = False
End Function

Public Sub PPTK_LibraryList(ByRef names() As String, ByRef paths() As String)
    Dim folder As String
    Dim found As String
    Dim count As Long
    Dim index As Long
    Dim spins As Long
    Dim holdName As String
    Dim holdPath As String
    folder = PPTK_LibraryFolder()
    count = 0
    ReDim names(1 To 1)
    ReDim paths(1 To 1)
    If Not PPTK_LibFolderExists(folder) Then
        ReDim names(0 To 0)
        ReDim paths(0 To 0)
        Exit Sub
    End If
    found = Dir$(folder & "*.pptx")
    Do While Len(found) > 0
        If Left$(found, 1) <> "~" And LCase$(Right$(found, 5)) = ".pptx" Then
            count = count + 1
            ReDim Preserve names(1 To count)
            ReDim Preserve paths(1 To count)
            names(count) = Left$(found, Len(found) - 5)
            paths(count) = folder & found
        End If
        found = Dir$()
    Loop
    If count < 1 Then
        ReDim names(0 To 0)
        ReDim paths(0 To 0)
        Exit Sub
    End If
    For index = 1 To count - 1
        For spins = index + 1 To count
            If StrComp(names(spins), names(index), vbTextCompare) < 0 Then
                holdName = names(index)
                holdPath = paths(index)
                names(index) = names(spins)
                paths(index) = paths(spins)
                names(spins) = holdName
                paths(spins) = holdPath
            End If
        Next spins
    Next index
End Sub

Private Function PPTK_LibHasSlideSelection() As Boolean
    If Application.Presentations.Count < 1 Then Exit Function
    If ActiveWindow Is Nothing Then Exit Function
    If ActiveWindow.Selection.Type <> ppSelectionSlides Then Exit Function
    If ActiveWindow.Selection.SlideRange Is Nothing Then Exit Function
    PPTK_LibHasSlideSelection = (ActiveWindow.Selection.SlideRange.Count > 0)
End Function

Private Function PPTK_LibInsertAfter(ByVal sourceWindow As DocumentWindow, ByVal source As Presentation) As Long
    Dim selected As SlideRange
    Dim index As Long
    Dim highest As Long
    On Error Resume Next
    If source.Slides.Count < 1 Then
        PPTK_LibInsertAfter = 0
        Exit Function
    End If
    highest = 0
    If sourceWindow.Selection.Type = ppSelectionSlides Then
        Set selected = sourceWindow.Selection.SlideRange
        For index = 1 To selected.Count
            If selected(index).SlideIndex > highest Then highest = selected(index).SlideIndex
        Next index
    End If
    If highest < 1 Then highest = sourceWindow.View.Slide.SlideIndex
    If highest < 1 Then highest = source.Slides.Count
    PPTK_LibInsertAfter = highest
End Function

Private Function PPTK_LibDefaultFolder() As String
    Dim profile As String
    profile = Environ$("USERPROFILE")
    If Len(profile) = 0 Then
        PPTK_LibDefaultFolder = CurDir$
    Else
        PPTK_LibDefaultFolder = profile & "\Documents\PPTK Slide Library"
    End If
End Function

Private Function PPTK_LibNormalizeFolder(ByVal folder As String) As String
    folder = Trim$(folder)
    If Len(folder) = 0 Then folder = PPTK_LibDefaultFolder()
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    PPTK_LibNormalizeFolder = folder
End Function

Private Function PPTK_LibFolderExists(ByVal folder As String) As Boolean
    Dim probe As String
    On Error Resume Next
    probe = folder
    If Right$(probe, 1) = "\" Then probe = Left$(probe, Len(probe) - 1)
    PPTK_LibFolderExists = (Len(probe) > 0 And Len(Dir$(probe, vbDirectory)) > 0)
End Function

Private Function PPTK_LibEnsureFolder(ByVal folder As String) As Boolean
    Dim parent As String
    On Error GoTo Failed
    If PPTK_LibFolderExists(folder) Then
        PPTK_LibEnsureFolder = True
        Exit Function
    End If
    parent = folder
    If Right$(parent, 1) = "\" Then parent = Left$(parent, Len(parent) - 1)
    MkDir parent
    PPTK_LibEnsureFolder = PPTK_LibFolderExists(folder)
    Exit Function
Failed:
    PPTK_LibEnsureFolder = False
End Function

Private Function PPTK_LibFolderWritable(ByVal folder As String) As Boolean
    Dim probe As String
    Dim handle As Integer
    On Error GoTo Failed
    probe = folder & "PPTK-library-write.tmp"
    handle = FreeFile
    Open probe For Output As #handle
    Close #handle
    Kill probe
    PPTK_LibFolderWritable = (Len(Dir$(probe)) = 0)
    Exit Function
Failed:
    On Error Resume Next
    Close #handle
    If Len(Dir$(probe)) > 0 Then Kill probe
    PPTK_LibFolderWritable = False
End Function

Private Function PPTK_LibUniquePath(ByVal folder As String, ByVal stem As String) As String
    Dim candidate As String
    Dim suffix As Long
    candidate = folder & stem & ".pptx"
    suffix = 2
    Do While Len(Dir$(candidate)) > 0
        candidate = folder & stem & " " & CStr(suffix) & ".pptx"
        suffix = suffix + 1
        If suffix > 100 Then Exit Function
    Loop
    PPTK_LibUniquePath = candidate
End Function

Private Function PPTK_LibPreviewPath(ByVal libraryPath As String) As String
    If LCase$(Right$(libraryPath, 5)) <> ".pptx" Then Exit Function
    PPTK_LibPreviewPath = Left$(libraryPath, Len(libraryPath) - 5) & ".jpg"
End Function

Private Function PPTK_LibTempPath() As String
    Dim folder As String
    folder = Environ$("TEMP")
    If Len(folder) = 0 Then folder = CurDir$
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    Randomize
    PPTK_LibTempPath = folder & "PPTK-slides-" & Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd() * 1000000)) & ".pptx"
End Function

Private Function PPTK_LibIsTempName(ByVal fileName As String) As Boolean
    PPTK_LibIsTempName = (StrComp(Left$(fileName, 12), "PPTK-slides-", vbTextCompare) = 0)
End Function

Private Sub PPTK_LibDeleteTemp(ByVal path As String)
    Dim slash As Long
    Dim leaf As String
    Dim attempt As Long
    If Len(path) = 0 Then Exit Sub
    slash = InStrRev(path, "\")
    If slash > 0 Then
        leaf = Mid$(path, slash + 1)
    Else
        leaf = path
    End If
    If Not PPTK_LibIsTempName(leaf) Then Exit Sub
    For attempt = 1 To 20
        On Error Resume Next
        If Len(Dir$(path)) = 0 Then Exit Sub
        Kill path
        If Err.Number = 0 Then Exit Sub
        Err.Clear
        DoEvents
    Next attempt
End Sub

Private Sub PPTK_LibCollectIndexes(ByVal selected As SlideRange, ByRef indexes() As Long)
    Dim index As Long
    ReDim indexes(1 To selected.Count)
    For index = 1 To selected.Count
        indexes(index) = selected(index).SlideIndex
    Next index
    PPTK_LibSortUnique indexes
End Sub

Private Sub PPTK_LibSortUnique(ByRef values() As Long)
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

Private Sub PPTK_LibKeepIndexes(ByVal copyPres As Presentation, ByRef indexes() As Long)
    Dim keep() As Boolean
    Dim index As Long
    Dim kept As Long
    If Not PPTK_LibIsTempName(copyPres.Name) Then Err.Raise 5
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

Private Sub PPTK_LibRestoreSelection(ByVal sourceWindow As DocumentWindow, ByVal source As Presentation, ByRef indexes() As Long)
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

Private Function PPTK_LibSameFile(ByVal source As Presentation, ByVal destination As String) As Boolean
    If Len(source.Path) = 0 Then Exit Function
    PPTK_LibSameFile = (StrComp(source.FullName, destination, vbTextCompare) = 0)
End Function

Private Function PPTK_LibSanitize(ByVal value As String) As String
    Dim invalidChars As String
    Dim index As Long
    invalidChars = "\/:*?""<>|"
    value = Replace$(value, vbCr, " ")
    value = Replace$(value, vbLf, " ")
    For index = 1 To Len(invalidChars)
        value = Replace$(value, Mid$(invalidChars, index, 1), " ")
    Next index
    Do While InStr(value, "  ") > 0
        value = Replace$(value, "  ", " ")
    Loop
    PPTK_LibSanitize = Trim$(value)
End Function

Private Function PPTK_LibFirstText(ByVal target As Slide) As String
    Dim shapeIndex As Long
    Dim shapeText As String
    On Error Resume Next
    For shapeIndex = 1 To target.Shapes.Count
        If target.Shapes(shapeIndex).HasTextFrame Then
            shapeText = Trim$(target.Shapes(shapeIndex).TextFrame.TextRange.Text)
            If Len(shapeText) > 0 Then
                PPTK_LibFirstText = shapeText
                Exit Function
            End If
        End If
    Next shapeIndex
End Function
