Attribute VB_Name = "ModulePPTKFlip"
Option Explicit

' Horizontal layout mirror for the current slide or every slide.
'
' Every object keeps its size, text, fonts, colors, and reading direction.
' Its box is mirrored across the slide width: newLeft = SlideWidth - Left - Width.
' Rotation becomes 360 - Rotation, so tilted objects lean the other way.
'
' Geometry that has a direction is mirrored with Shape.Flip. PowerPoint keeps
' text upright and readable in a horizontally flipped shape (checked by
' exporting a flipped chevron and arrow that contain Arabic and English text).
' Flip is used for AutoShapes, freeforms, lines, and unglued connectors.
' Shapes with a picture or texture fill are not flipped, because the fill
' image would be mirrored. Text boxes and placeholders are not flipped.
'
' Pictures, SVG icons, charts, media, and OLE objects are moved and their
' rotation is mirrored. Their content is not mirrored, so logos, screenshots,
' and chart text stay readable. Direction-dependent icons keep their direction.
'
' Text: left and right paragraph alignment are swapped. Centered, justified,
' and distributed paragraphs and the paragraph reading direction stay.
'
' Tables: the table direction is toggled (left-to-right / right-to-left),
' which reverses the visible column order without touching cell content.
' Cell alignment is swapped like other text.
'
' SmartArt: the SmartArt Reverse setting is toggled.
'
' Groups: the members are mirrored inside the group box, then the group box
' is mirrored on the slide. A rotated group is set to 0 degrees while its
' members move, then gets the mirrored rotation.
'
' Connectors glued to shapes are re-glued after the shapes move. A glue
' point on a flipped shape keeps its index, because Flip mirrors the sites.
' On an unflipped shape with 4 or 8 sites the site is mirrored
' (left <-> right). Other site layouts keep their index.
'
' Running the command twice restores the original layout (positions,
' rotation, flips, alignment, table direction, SmartArt direction).
'
' Recovery: each command is one Undo entry. Flip Entire Presentation also
' saves a backup copy of the presentation before it changes anything.
'
' Not mirrored: slide masters, layouts, notes, animation motion paths,
' chart internals, and members of nested groups that have their own rotation.

Private Const SITE_MAP_A As Long = 4
Private Const SITE_MAP_B As Long = 8

Private Type PPTKGlue
    Connector As Shape
    BeginShape As Shape
    BeginSite As Long
    EndShape As Shape
    EndSite As Long
    AxisSum As Single
    OldLeft As Single
    OldTop As Single
    OldWidth As Single
    OldHeight As Single
    OldFlipH As MsoTriState
    OldFlipV As MsoTriState
End Type

Private GlueList() As PPTKGlue
Private GlueCount As Long
Private GlueIds As Collection
Private FlippedIds As Collection
Private FailedCount As Long

Public Sub PPTK_FlipCurrentSlide()
    Dim sld As Slide
    Dim slideWidth As Single
    Dim reason As String

    If Not PPTK_TryCurrentSlide(sld, reason) Then
        MsgBox reason, vbExclamation, "PPTK"
        Exit Sub
    End If

    On Error GoTo Failed
    If sld.Shapes.Count = 0 Then
        MsgBox "Nothing to flip.", vbExclamation, "PPTK"
        Exit Sub
    End If
    If Not PPTK_ReadSlideWidth(sld.Parent, slideWidth) Then
        MsgBox "Could not read the slide size.", vbExclamation, "PPTK"
        Exit Sub
    End If

    PPTK_StartUndo "Flip slide"
    PPTK_FlipSlide sld, slideWidth
    If FailedCount > 0 Then
        MsgBox FailedCount & " object(s) could not be mirrored. Press Ctrl+Z to undo the flip.", vbExclamation, "PPTK"
    End If
    Exit Sub
Failed:
    MsgBox "Could not flip the current slide. Press Ctrl+Z to undo any partial change.", vbExclamation, "PPTK"
End Sub

Public Sub PPTK_FlipEntirePresentation()
    Dim pres As Presentation
    Dim sld As Slide
    Dim slideWidth As Single
    Dim backupPath As String
    Dim answer As VbMsgBoxResult
    Dim totalFailed As Long
    Dim slideTotal As Long

    On Error Resume Next
    Err.Clear
    Set pres = ActivePresentation
    If Err.Number <> 0 Or pres Is Nothing Then
        MsgBox "No presentation is open.", vbExclamation, "PPTK"
        Exit Sub
    End If

    On Error GoTo Failed
    slideTotal = pres.Slides.Count
    If slideTotal = 0 Or PPTK_CountShapes(pres) = 0 Then
        MsgBox "Nothing to flip.", vbExclamation, "PPTK"
        Exit Sub
    End If
    If Not PPTK_ReadSlideWidth(pres, slideWidth) Then
        MsgBox "Could not read the slide size.", vbExclamation, "PPTK"
        Exit Sub
    End If

    answer = MsgBox("Mirror the layout of all " & slideTotal & " slides horizontally?" & vbCr & vbCr & _
        "A backup copy of the presentation is saved first. Masters, layouts, and notes are not changed.", _
        vbOKCancel + vbQuestion + vbDefaultButton2, "PPTK - Flip Entire Presentation")
    If answer <> vbOK Then Exit Sub

    backupPath = PPTK_SaveBackupCopy(pres, "PPTK Flip Backups", "before flip")
    If Len(backupPath) = 0 Then
        answer = MsgBox("A backup copy could not be saved." & vbCr & _
            "Flip anyway? You can still undo with Ctrl+Z.", vbYesNo + vbExclamation + vbDefaultButton2, "PPTK")
        If answer <> vbYes Then Exit Sub
    End If

    PPTK_StartUndo "Flip presentation"
    totalFailed = 0
    For Each sld In pres.Slides
        PPTK_FlipSlide sld, slideWidth
        totalFailed = totalFailed + FailedCount
    Next sld

    If totalFailed > 0 Then
        MsgBox "The presentation was flipped, but " & totalFailed & " object(s) could not be mirrored." & vbCr & _
            PPTK_BackupNote(backupPath), vbExclamation, "PPTK"
    ElseIf Len(backupPath) > 0 Then
        MsgBox "The presentation was flipped." & vbCr & PPTK_BackupNote(backupPath), vbInformation, "PPTK"
    End If
    Exit Sub
Failed:
    MsgBox "Could not flip the presentation. Press Ctrl+Z to undo any partial change." & vbCr & _
        PPTK_BackupNote(backupPath), vbExclamation, "PPTK"
End Sub

Private Function PPTK_BackupNote(ByVal backupPath As String) As String
    If Len(backupPath) = 0 Then
        PPTK_BackupNote = "No backup copy was saved."
    Else
        PPTK_BackupNote = "Backup copy: " & backupPath
    End If
End Function

' Saves a copy of the presentation (its current, unsaved state included) to
' Documents\<folderName> (the profile's Documents, else OneDrive\Documents when
' Known Folder Move is on, else %TEMP%), named "<name> (<label> yyyy-mm-dd hhnnss).<ext>".
' An existing file is never overwritten: a counter is added. File checks use
' Scripting.FileSystemObject, which handles Unicode (for example Arabic) names.
' Returns the path, or "" when no copy was made.
' Shared by Flip Entire Presentation and Find & Replace Colors.
Public Function PPTK_SaveBackupCopy(ByVal pres As Presentation, ByVal folderName As String, ByVal label As String) As String
    Dim fso As Object
    Dim folder As String
    Dim baseName As String
    Dim extension As String
    Dim dotAt As Long
    Dim target As String
    Dim index As Long
    Dim stem As String
    Dim attempt As Long

    On Error GoTo Clean
    PPTK_SaveBackupCopy = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Documents, not %TEMP%: Windows Storage Sense and Disk Cleanup delete temp files.
    folder = ""
    If Len(Environ$("USERPROFILE")) > 0 Then
        If fso.FolderExists(Environ$("USERPROFILE") & "\Documents") Then folder = Environ$("USERPROFILE") & "\Documents"
    End If
    If Len(folder) = 0 And Len(Environ$("OneDrive")) > 0 Then
        If fso.FolderExists(Environ$("OneDrive") & "\Documents") Then folder = Environ$("OneDrive") & "\Documents"
    End If
    If Len(folder) = 0 Then folder = Environ$("TEMP")
    If Len(folder) = 0 Then Exit Function
    folder = folder & "\" & folderName
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder

    baseName = pres.Name
    dotAt = InStrRev(baseName, ".")
    If dotAt > 1 Then
        extension = Mid$(baseName, dotAt)
        baseName = Left$(baseName, dotAt - 1)
    Else
        extension = ".pptx"
    End If
    For index = 1 To Len(baseName)
        If InStr("\/:*?""<>|", Mid$(baseName, index, 1)) > 0 Then Mid$(baseName, index, 1) = "_"
    Next index

    stem = folder & "\" & baseName & " (" & label & " " & Format$(Now, "yyyy-mm-dd hhnnss") & ")"
    target = stem & extension
    attempt = 1
    Do While fso.FileExists(target)
        attempt = attempt + 1
        target = stem & " " & attempt & extension
    Loop
    pres.SaveCopyAs target
    If Not fso.FileExists(target) Then Exit Function
    PPTK_SaveBackupCopy = target
    Exit Function
Clean:
    PPTK_SaveBackupCopy = ""
    Err.Clear
End Function

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

Private Function PPTK_ReadSlideWidth(ByVal pres As Presentation, ByRef slideWidth As Single) As Boolean
    On Error Resume Next
    Err.Clear
    slideWidth = pres.PageSetup.SlideWidth
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
    Application.StartNewUndoEntry
End Sub

' One slide: record glued connectors, mirror everything else, then re-glue.
Private Sub PPTK_FlipSlide(ByVal sld As Slide, ByVal slideWidth As Single)
    Dim index As Long
    Dim total As Long
    Dim shp As Shape

    On Error Resume Next
    FailedCount = 0
    GlueCount = 0
    Erase GlueList
    Set GlueIds = New Collection
    Set FlippedIds = New Collection

    total = sld.Shapes.Count
    For index = 1 To total
        Err.Clear
        Set shp = sld.Shapes(index)
        If Err.Number = 0 Then PPTK_RecordGlue shp, slideWidth
    Next index

    For index = 1 To total
        Err.Clear
        Set shp = sld.Shapes(index)
        If Err.Number = 0 Then
            PPTK_MirrorShape shp, slideWidth
        Else
            FailedCount = FailedCount + 1
        End If
    Next index

    For index = 1 To GlueCount
        PPTK_Reglue index
    Next index

    Set GlueIds = Nothing
    Set FlippedIds = Nothing
    Erase GlueList
    GlueCount = 0
End Sub

' Glued connectors are captured before anything moves, because PowerPoint
' drags a glued end along with its shape.
Private Sub PPTK_RecordGlue(ByVal shp As Shape, ByVal axisSum As Single)
    Dim kind As MsoShapeType
    Dim index As Long
    Dim childCount As Long
    Dim groupAxis As Single

    On Error Resume Next
    Err.Clear
    kind = shp.Type
    If Err.Number <> 0 Then Exit Sub

    If kind = msoGroup Then
        Err.Clear
        groupAxis = 2 * shp.Left + shp.Width
        childCount = shp.GroupItems.Count
        If Err.Number <> 0 Then Exit Sub
        For index = 1 To childCount
            PPTK_RecordGlue shp.GroupItems(index), groupAxis
        Next index
        Exit Sub
    End If

    If Not PPTK_IsConnector(shp) Then Exit Sub

    Dim glue As PPTKGlue
    Dim beginOn As Boolean
    Dim endOn As Boolean

    Err.Clear
    beginOn = (shp.ConnectorFormat.BeginConnected = msoTrue)
    If Err.Number <> 0 Then beginOn = False
    Err.Clear
    endOn = (shp.ConnectorFormat.EndConnected = msoTrue)
    If Err.Number <> 0 Then endOn = False
    Err.Clear
    If Not beginOn And Not endOn Then Exit Sub

    Set glue.Connector = shp
    If beginOn Then
        Set glue.BeginShape = shp.ConnectorFormat.BeginConnectedShape
        glue.BeginSite = shp.ConnectorFormat.BeginConnectionSite
    End If
    If endOn Then
        Set glue.EndShape = shp.ConnectorFormat.EndConnectedShape
        glue.EndSite = shp.ConnectorFormat.EndConnectionSite
    End If
    glue.AxisSum = axisSum
    glue.OldLeft = shp.Left
    glue.OldTop = shp.Top
    glue.OldWidth = shp.Width
    glue.OldHeight = shp.Height
    glue.OldFlipH = shp.HorizontalFlip
    glue.OldFlipV = shp.VerticalFlip
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    GlueCount = GlueCount + 1
    ReDim Preserve GlueList(1 To GlueCount)
    GlueList(GlueCount) = glue
    GlueIds.Add True, CStr(shp.Id)
    Err.Clear
End Sub

Private Sub PPTK_MirrorShape(ByVal shp As Shape, ByVal axisSum As Single)
    Dim kind As MsoShapeType
    Dim newLeft As Single
    Dim isTable As Boolean
    Dim isChart As Boolean
    Dim isSmartArt As Boolean

    On Error Resume Next
    If shp Is Nothing Then Exit Sub
    If PPTK_InSet(GlueIds, shp.Id) Then Exit Sub

    Err.Clear
    kind = shp.Type
    If Err.Number <> 0 Then
        FailedCount = FailedCount + 1
        Exit Sub
    End If

    If kind = msoGroup Then
        PPTK_MirrorGroup shp, axisSum
        Exit Sub
    End If

    Err.Clear
    newLeft = axisSum - shp.Left - shp.Width
    If Err.Number <> 0 Then
        FailedCount = FailedCount + 1
        Exit Sub
    End If
    Err.Clear
    shp.Left = newLeft
    If Err.Number <> 0 Then
        FailedCount = FailedCount + 1
        Exit Sub
    End If

    Err.Clear
    isTable = shp.HasTable
    If Err.Number <> 0 Then isTable = False
    If isTable Then
        PPTK_MirrorTable shp
        Exit Sub
    End If

    Err.Clear
    isSmartArt = shp.HasSmartArt
    If Err.Number <> 0 Then isSmartArt = False
    If isSmartArt Then
        PPTK_MirrorSmartArt shp
        Exit Sub
    End If

    Err.Clear
    isChart = shp.HasChart
    If Err.Number <> 0 Then isChart = False
    If isChart Then Exit Sub

    ' Shape.Flip already turns Rotation into 360 - Rotation (checked in
    ' PowerPoint: a 15 degree arrow becomes 345 degrees when flipped), so the
    ' rotation is mirrored by hand only for shapes that are not flipped.
    If PPTK_FlipsGeometry(shp, kind) Then
        Err.Clear
        shp.Flip msoFlipHorizontal
        If Err.Number = 0 Then
            PPTK_AddToSet FlippedIds, shp.Id
        Else
            Err.Clear
            PPTK_MirrorRotation shp
        End If
    Else
        PPTK_MirrorRotation shp
    End If

    If PPTK_HasText(shp) Then PPTK_SwapShapeAlignment shp
    Err.Clear
End Sub

Private Sub PPTK_MirrorGroup(ByVal grp As Shape, ByVal axisSum As Single)
    Dim oldLeft As Single
    Dim oldWidth As Single
    Dim angle As Single
    Dim innerAxis As Single
    Dim index As Long
    Dim childCount As Long

    On Error Resume Next
    Err.Clear
    angle = grp.Rotation
    If Err.Number <> 0 Then angle = 0
    Err.Clear
    If angle <> 0 Then grp.Rotation = 0

    oldLeft = grp.Left
    oldWidth = grp.Width
    childCount = grp.GroupItems.Count
    If Err.Number <> 0 Then
        Err.Clear
        If angle <> 0 Then grp.Rotation = angle
        FailedCount = FailedCount + 1
        Exit Sub
    End If

    innerAxis = 2 * oldLeft + oldWidth
    For index = 1 To childCount
        Err.Clear
        PPTK_MirrorShape grp.GroupItems(index), innerAxis
    Next index

    Err.Clear
    grp.Left = axisSum - oldLeft - oldWidth
    If Err.Number <> 0 Then FailedCount = FailedCount + 1
    Err.Clear
    If angle <> 0 Then grp.Rotation = PPTK_MirroredAngle(angle)
    Err.Clear
End Sub

Private Function PPTK_MirroredAngle(ByVal angle As Single) As Single
    Dim result As Single
    result = 360 - angle
    Do While result >= 360
        result = result - 360
    Loop
    Do While result < 0
        result = result + 360
    Loop
    PPTK_MirroredAngle = result
End Function

Private Sub PPTK_MirrorRotation(ByVal shp As Shape)
    Dim angle As Single
    On Error Resume Next
    Err.Clear
    angle = shp.Rotation
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    If angle <> 0 Then shp.Rotation = PPTK_MirroredAngle(angle)
    Err.Clear
End Sub

' True when Shape.Flip mirrors the outline without mirroring visible content.
Private Function PPTK_FlipsGeometry(ByVal shp As Shape, ByVal kind As MsoShapeType) As Boolean
    Dim fillKind As MsoFillType
    PPTK_FlipsGeometry = False
    On Error Resume Next

    If PPTK_IsConnector(shp) Then
        PPTK_FlipsGeometry = True
        Exit Function
    End If

    Select Case kind
        Case msoLine
            PPTK_FlipsGeometry = True
        Case msoAutoShape, msoFreeform
            Err.Clear
            fillKind = shp.Fill.Type
            If Err.Number <> 0 Then
                Err.Clear
                Exit Function
            End If
            If fillKind = msoFillPicture Or fillKind = msoFillTextured Then Exit Function
            PPTK_FlipsGeometry = True
    End Select
    Err.Clear
End Function

Private Function PPTK_IsConnector(ByVal shp As Shape) As Boolean
    On Error Resume Next
    PPTK_IsConnector = (shp.Connector = msoTrue)
    If Err.Number <> 0 Then PPTK_IsConnector = False
    Err.Clear
End Function

Private Sub PPTK_Reglue(ByVal index As Long)
    Dim glue As PPTKGlue
    Dim k As Shape
    Dim wantH As MsoTriState

    On Error Resume Next
    glue = GlueList(index)
    Set k = glue.Connector
    If k Is Nothing Then Exit Sub

    If Not glue.BeginShape Is Nothing And Not glue.EndShape Is Nothing Then
        Err.Clear
        k.ConnectorFormat.BeginConnect glue.BeginShape, PPTK_MirrorSite(glue.BeginShape, glue.BeginSite)
        k.ConnectorFormat.EndConnect glue.EndShape, PPTK_MirrorSite(glue.EndShape, glue.EndSite)
        If Err.Number <> 0 Then FailedCount = FailedCount + 1
        Err.Clear
        Exit Sub
    End If

    ' One end glued: put the free end where the mirror says, then glue the other end.
    Err.Clear
    If Not glue.BeginShape Is Nothing Then k.ConnectorFormat.BeginDisconnect
    If Not glue.EndShape Is Nothing Then k.ConnectorFormat.EndDisconnect
    Err.Clear
    If glue.OldFlipH = msoTrue Then wantH = msoFalse Else wantH = msoTrue
    k.Width = glue.OldWidth
    k.Height = glue.OldHeight
    k.Left = glue.AxisSum - glue.OldLeft - glue.OldWidth
    k.Top = glue.OldTop
    If k.HorizontalFlip <> wantH Then k.Flip msoFlipHorizontal
    If k.VerticalFlip <> glue.OldFlipV Then k.Flip msoFlipVertical
    If Not glue.BeginShape Is Nothing Then
        k.ConnectorFormat.BeginConnect glue.BeginShape, PPTK_MirrorSite(glue.BeginShape, glue.BeginSite)
    End If
    If Not glue.EndShape Is Nothing Then
        k.ConnectorFormat.EndConnect glue.EndShape, PPTK_MirrorSite(glue.EndShape, glue.EndSite)
    End If
    If Err.Number <> 0 Then FailedCount = FailedCount + 1
    Err.Clear
End Sub

' Connection sites run counter-clockwise from the top. Flip mirrors the sites
' of a flipped shape, so its index stays. Unflipped 4- and 8-site shapes
' (rectrotations, text boxes, placeholders, pictures, ovals) swap left and right.
Private Function PPTK_MirrorSite(ByVal target As Shape, ByVal site As Long) As Long
    Dim siteCount As Long
    On Error Resume Next
    PPTK_MirrorSite = site
    If PPTK_InSet(FlippedIds, target.Id) Then Exit Function
    Err.Clear
    siteCount = target.ConnectionSiteCount
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    If site < 1 Or site > siteCount Then Exit Function
    If siteCount = SITE_MAP_A Or siteCount = SITE_MAP_B Then
        PPTK_MirrorSite = ((siteCount - (site - 1)) Mod siteCount) + 1
    End If
End Function

Private Function PPTK_InSet(ByVal items As Collection, ByVal shapeId As Long) As Boolean
    Dim probe As Variant
    PPTK_InSet = False
    If items Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    probe = items(CStr(shapeId))
    PPTK_InSet = (Err.Number = 0)
    Err.Clear
End Function

Private Sub PPTK_AddToSet(ByVal items As Collection, ByVal shapeId As Long)
    On Error Resume Next
    If items Is Nothing Then Exit Sub
    items.Add True, CStr(shapeId)
    Err.Clear
End Sub

Private Function PPTK_HasText(ByVal shp As Shape) As Boolean
    Dim hasFrame As MsoTriState
    Dim hasBody As MsoTriState

    On Error Resume Next
    PPTK_HasText = False

    Err.Clear
    hasFrame = shp.HasTextFrame
    If Err.Number <> 0 Or hasFrame <> msoTrue Then
        Err.Clear
        Exit Function
    End If

    Err.Clear
    hasBody = shp.TextFrame2.HasText
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    PPTK_HasText = (hasBody = msoTrue)
End Function

Private Sub PPTK_MirrorTable(ByVal shp As Shape)
    Dim rowCount As Long
    Dim colCount As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim cellShape As Shape
    Dim tableDir As PpDirection

    On Error Resume Next
    Err.Clear
    tableDir = shp.Table.TableDirection
    If Err.Number = 0 Then
        If tableDir = ppDirectionRightToLeft Then
            shp.Table.TableDirection = ppDirectionLeftToRight
        Else
            shp.Table.TableDirection = ppDirectionRightToLeft
        End If
    End If
    If Err.Number <> 0 Then FailedCount = FailedCount + 1

    Err.Clear
    rowCount = shp.Table.Rows.Count
    colCount = shp.Table.Columns.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    For rowIndex = 1 To rowCount
        For colIndex = 1 To colCount
            Err.Clear
            Set cellShape = Nothing
            Set cellShape = shp.Table.Cell(rowIndex, colIndex).Shape
            If Err.Number = 0 Then
                If Not cellShape Is Nothing Then
                    If PPTK_HasText(cellShape) Then PPTK_SwapShapeAlignment cellShape
                End If
            End If
        Next colIndex
    Next rowIndex
    Err.Clear
End Sub

Private Sub PPTK_MirrorSmartArt(ByVal shp As Shape)
    On Error Resume Next
    Err.Clear
    If shp.SmartArt.Reverse = msoTrue Then
        shp.SmartArt.Reverse = msoFalse
    Else
        shp.SmartArt.Reverse = msoTrue
    End If
    If Err.Number <> 0 Then FailedCount = FailedCount + 1
    Err.Clear
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
    Dim align As MsoParagraphAlignment

    On Error Resume Next
    If textRange Is Nothing Then Exit Sub

    Err.Clear
    align = textRange.ParagraphFormat.Alignment
    If Err.Number <> 0 Then Exit Sub
    If align = msoAlignLeft Then
        textRange.ParagraphFormat.Alignment = msoAlignRight
    ElseIf align = msoAlignRight Then
        textRange.ParagraphFormat.Alignment = msoAlignLeft
    End If
End Sub
