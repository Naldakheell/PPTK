Attribute VB_Name = "ModulePPTKGrid"
Option Explicit

' PPTK row and column commands. Keep bounds and expand/shrink are separate
' implementations. They share recognition only.
'
' Keep bounds (PPTK_AddRightColumn, PPTK_AddBottomRow,
' PPTK_DeleteRightColumn, PPTK_DeleteBottomRow) leaves the original outer
' rectangle in place. The bounds are captured before any change. Add
' duplicates the outer edge at its current size (formatting only: the new
' cells are empty, existing text is untouched), then resizes flexible
' shapes and gaps on that axis so every column or row fits between the
' saved edges. Delete removes the outer edge and expands the survivors
' back to those same edges. The other axis is not resized, so a column
' change does not move the vertical footprint and a row change does not
' move the horizontal footprint.
'
' Flexible sizes and gaps share one scale. Each column or row also keeps
' a fixed floor: the widest or tallest shape on that axis that must not
' be resized. Pictures, linked pictures, locked aspect ratios, groups
' that contain either are fixed. A TextFrame2 shape-to-fit-text shape is
' fixed in height; its width is fixed only when the text does not wrap.
' Fixed shapes are only moved inside their slot. If the fixed floors
' cannot fit in the original span with zero gaps, nothing is duplicated
' or deleted. A new gap copies the gap before the duplicated edge, or
' zero when the grid has only one column or row. Negative gaps are
' treated as zero so the result does not overlap.
'
' A native table is keep-bounds only when it is the single selected shape
' and it has no merged cells. Merged cells abort with a message. Otherwise
' the new or remaining row heights or column widths are scaled to the
' saved shape height or width, and Left and Top are restored. The last
' row or column absorbs rounding. The added table row or column is empty.
'
' Expand / shrink (PPTK_AddRightColumnExpand, PPTK_AddBottomRowExpand,
' PPTK_DeleteRightColumnShrink, PPTK_DeleteBottomRowShrink) is the outward
' geometry. Add duplicates the outer edge (formatting only, empty text)
' beyond the grid by the median measured gap on that axis (zero for a
' touching grid), else the other axis's gap, else 8 points, and
' PPTK_ShiftInside pulls the selection back
' only when the copies would leave the slide. Delete removes the outer
' edge and leaves the other shapes where they are. A native table uses
' Columns.Add or Rows.Add, or deletes the last row or column, so the
' table's outer size changes. A one-row or one-column table is not deleted.
'
' Recognition is the same for both modes. Shape grids use the current
' shape selection only. Nothing else on the slide is moved or deleted.
' Pictures, lines, connectors, media, OLE, ink, and charts are ignored.
' A group counts as one shape. Extra tables are ignored.
'
' Rows share a vertical center within Max(6, Min(18, 0.35 * median height))
' points, or overlap vertically by at least half of the shorter shape, so
' top- or bottom-aligned cells of different heights form one row. Each row's
' shapes must share one vertical band (within that tolerance); a shape that
' spans two rows makes the selection ambiguous and it is rejected. Columns share a horizontal center within the same rule using width.
' A grid is recognized only when every row has the same count and the columns
' line up within that tolerance. A single row or a single column is valid.
' A ragged or staggered selection does nothing, including deletion, and
' says why. A one-column grid is not deleted. A one-row grid is not deleted.
'
' Data safety. Every command is one Undo step. Nothing is deleted until the
' plan has been validated. Shape edges are deleted through hidden recovery
' copies: if any delete fails, the deleted shapes come back from their
' copies with their original position, name, and stacking order. A keep-
' bounds table delete makes a hidden copy of the whole table first. If the
' table cannot keep its size after the edge is gone, the copy replaces it,
' so cell content in the deleted edge is never lost. A legacy table delete
' is one Rows/Columns.Delete call, which either completes or raises first.

Private Const GapDefault As Single = 8
Private Const SlideMargin As Single = 8
Private Const MinAxisSize As Single = 1
Private Const BoundSlack As Single = 1.25

Public Sub PPTK_AddRightColumn()
    PPTK_StartUndo
    PPTK_AddAxisKeep False
End Sub

Public Sub PPTK_AddBottomRow()
    PPTK_StartUndo
    PPTK_AddAxisKeep True
End Sub

Public Sub PPTK_DeleteRightColumn()
    PPTK_StartUndo
    PPTK_DeleteAxisKeep False
End Sub

Public Sub PPTK_DeleteBottomRow()
    PPTK_StartUndo
    PPTK_DeleteAxisKeep True
End Sub

Public Sub PPTK_AddRightColumnExpand()
    PPTK_StartUndo
    PPTK_AddAxisExpand False
End Sub

Public Sub PPTK_AddBottomRowExpand()
    PPTK_StartUndo
    PPTK_AddAxisExpand True
End Sub

Public Sub PPTK_DeleteRightColumnShrink()
    PPTK_StartUndo
    PPTK_DeleteAxisShrink False
End Sub

Public Sub PPTK_DeleteBottomRowShrink()
    PPTK_StartUndo
    PPTK_DeleteAxisShrink True
End Sub

' Each command is one Undo step, so Ctrl+Z reverts the whole change.
Private Sub PPTK_StartUndo()
    On Error Resume Next
    Application.StartNewUndoEntry
    Err.Clear
End Sub

Private Function PPTK_NotGridMessage() As String
    PPTK_NotGridMessage = "Select the shapes of one grid (the same number of shapes in every row), or one table."
End Function

Private Sub PPTK_AddAxisKeep(ByVal bottomRow As Boolean)
    Dim tableShape As Shape

    If PPTK_SingleTable(tableShape) Then
        PPTK_KeepAddTable tableShape, bottomRow
        Exit Sub
    End If
    PPTK_KeepAddShapes bottomRow
End Sub

Private Sub PPTK_DeleteAxisKeep(ByVal bottomRow As Boolean)
    Dim tableShape As Shape

    If PPTK_SingleTable(tableShape) Then
        PPTK_KeepDeleteTable tableShape, bottomRow
        Exit Sub
    End If
    PPTK_KeepDeleteShapes bottomRow
End Sub

Private Sub PPTK_AddAxisExpand(ByVal bottomRow As Boolean)
    Dim tableShape As Shape
    Dim items() As Shape
    Dim count As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim rowOf() As Long
    Dim colOf() As Long
    Dim edge() As Shape
    Dim edgeCount As Long
    Dim copies() As Shape
    Dim index As Long
    Dim gap As Single
    Dim edgePos As Single
    Dim made As Long
    Dim lefts() As Single
    Dim tops() As Single
    Dim widths() As Single
    Dim heights() As Single
    Dim axisFixed() As Boolean

    If PPTK_SingleTable(tableShape) Then
        On Error GoTo TableFailed
        If bottomRow Then
            tableShape.Table.Rows.Add
        Else
            tableShape.Table.Columns.Add
        End If
        Exit Sub
TableFailed:
        PPTK_Tell "The table could not be extended."
        Exit Sub
    End If

    If Not PPTK_Recognize(items, count, rowCount, colCount, rowOf, colOf) Then
        PPTK_Tell PPTK_NotGridMessage()
        Exit Sub
    End If
    If Not PPTK_CaptureGrid(items, count, bottomRow, lefts, tops, widths, heights, axisFixed) Then
        PPTK_Tell "The grid could not be measured. The layout was not changed."
        Exit Sub
    End If
    On Error GoTo ExpandFailed
    If bottomRow Then
        edgeCount = colCount
        ReDim edge(1 To edgeCount)
        For index = 1 To count
            If rowOf(index) = rowCount Then Set edge(colOf(index)) = items(index)
        Next index
        gap = PPTK_PositiveGap(items, count, rowOf, colOf, rowCount, colCount, True)
        edgePos = edge(1).Top + edge(1).Height
        For index = 2 To edgeCount
            If edge(index).Top + edge(index).Height > edgePos Then edgePos = edge(index).Top + edge(index).Height
        Next index
    Else
        edgeCount = rowCount
        ReDim edge(1 To edgeCount)
        For index = 1 To count
            If colOf(index) = colCount Then Set edge(rowOf(index)) = items(index)
        Next index
        gap = PPTK_PositiveGap(items, count, rowOf, colOf, rowCount, colCount, False)
        edgePos = edge(1).Left + edge(1).Width
        For index = 2 To edgeCount
            If edge(index).Left + edge(index).Width > edgePos Then edgePos = edge(index).Left + edge(index).Width
        Next index
    End If

    ReDim copies(1 To edgeCount)
    made = 0
    For index = 1 To edgeCount
        If edge(index) Is Nothing Then GoTo ExpandUndo
        Set copies(index) = PPTK_NewCell(edge(index))
        If copies(index) Is Nothing Then GoTo ExpandUndo
        made = index
        If bottomRow Then
            copies(index).Left = edge(index).Left
            copies(index).Top = edgePos + gap
        Else
            copies(index).Top = edge(index).Top
            copies(index).Left = edgePos + gap
        End If
    Next index
    PPTK_ShiftInside items, count, copies, edgeCount, bottomRow
    On Error Resume Next
    items(1).Select
    For index = 2 To count
        items(index).Select False
    Next index
    For index = 1 To edgeCount
        copies(index).Select False
    Next index
    Exit Sub
ExpandFailed:
    Resume ExpandUndo
ExpandUndo:
    PPTK_RestoreShapes items, count, lefts, tops, widths, heights
    PPTK_DeleteCopies copies, made
    PPTK_Tell "The row or column could not be added. The layout was restored."
End Sub

Private Sub PPTK_DeleteAxisShrink(ByVal bottomRow As Boolean)
    Dim tableShape As Shape
    Dim items() As Shape
    Dim count As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim rowOf() As Long
    Dim colOf() As Long
    Dim index As Long
    Dim edge() As Shape
    Dim edgeCount As Long
    Dim reason As String

    If PPTK_SingleTable(tableShape) Then
        ' One Rows/Columns.Delete call either removes the edge or raises before it does.
        On Error GoTo TableFailed
        If bottomRow Then
            If tableShape.Table.Rows.Count <= 1 Then
                PPTK_Tell "The last row of a table is not deleted."
                Exit Sub
            End If
            tableShape.Table.Rows(tableShape.Table.Rows.Count).Delete
        Else
            If tableShape.Table.Columns.Count <= 1 Then
                PPTK_Tell "The last column of a table is not deleted."
                Exit Sub
            End If
            tableShape.Table.Columns(tableShape.Table.Columns.Count).Delete
        End If
        Exit Sub
TableFailed:
        PPTK_Tell "The table edge could not be deleted."
        Exit Sub
    End If

    If Not PPTK_Recognize(items, count, rowCount, colCount, rowOf, colOf) Then
        PPTK_Tell PPTK_NotGridMessage()
        Exit Sub
    End If
    If Not PPTK_LastEdgeGuard(bottomRow, rowCount, colCount) Then Exit Sub
    PPTK_CollectEdge items, count, rowOf, colOf, rowCount, colCount, bottomRow, edge, edgeCount
    If Not PPTK_DeleteEdgeSafely(edge, edgeCount, reason) Then
        PPTK_Tell reason
        Exit Sub
    End If
    PPTK_SelectSurvivors items, count, rowOf, colOf, rowCount, colCount, bottomRow
End Sub

' The final row or column is never deleted.
Private Function PPTK_LastEdgeGuard(ByVal bottomRow As Boolean, ByVal rowCount As Long, ByVal colCount As Long) As Boolean
    PPTK_LastEdgeGuard = True
    If bottomRow And rowCount <= 1 Then
        PPTK_Tell "The grid has one row. The last row is not deleted."
        PPTK_LastEdgeGuard = False
    ElseIf Not bottomRow And colCount <= 1 Then
        PPTK_Tell "The grid has one column. The last column is not deleted."
        PPTK_LastEdgeGuard = False
    End If
End Function

Private Sub PPTK_CollectEdge(ByRef items() As Shape, ByVal count As Long, ByRef rowOf() As Long, ByRef colOf() As Long, ByVal rowCount As Long, ByVal colCount As Long, ByVal bottomRow As Boolean, ByRef edge() As Shape, ByRef edgeCount As Long)
    Dim index As Long
    edgeCount = 0
    ReDim edge(1 To count)
    For index = 1 To count
        If (bottomRow And rowOf(index) = rowCount) Or (Not bottomRow And colOf(index) = colCount) Then
            edgeCount = edgeCount + 1
            Set edge(edgeCount) = items(index)
        End If
    Next index
End Sub

Private Sub PPTK_SelectSurvivors(ByRef items() As Shape, ByVal count As Long, ByRef rowOf() As Long, ByRef colOf() As Long, ByVal rowCount As Long, ByVal colCount As Long, ByVal bottomRow As Boolean)
    Dim index As Long
    Dim first As Boolean
    On Error Resume Next
    first = True
    For index = 1 To count
        If Not ((bottomRow And rowOf(index) = rowCount) Or (Not bottomRow And colOf(index) = colCount)) Then
            If first Then
                items(index).Select
                first = False
            Else
                items(index).Select False
            End If
        End If
    Next index
    Err.Clear
End Sub

' Deletes the edge shapes only after every one of them has a hidden recovery
' copy. If any delete fails, each deleted shape is replaced by its copy at its
' original position, name, and stacking order, and nothing is lost. The
' copies are removed after a successful delete.
Private Function PPTK_DeleteEdgeSafely(ByRef edge() As Shape, ByVal edgeCount As Long, ByRef failReason As String) As Boolean
    Dim backups() As Shape
    Dim names() As String
    Dim stack() As Long
    Dim oldLeft() As Single
    Dim oldTop() As Single
    Dim oldWidth() As Single
    Dim oldHeight() As Single
    Dim deleted() As Boolean
    Dim made As Long
    Dim index As Long
    Dim copyShape As Shape

    PPTK_DeleteEdgeSafely = False
    failReason = "The edge could not be deleted. Nothing was changed."
    If edgeCount < 1 Then Exit Function
    ReDim backups(1 To edgeCount)
    ReDim names(1 To edgeCount)
    ReDim stack(1 To edgeCount)
    ReDim oldLeft(1 To edgeCount)
    ReDim oldTop(1 To edgeCount)
    ReDim oldWidth(1 To edgeCount)
    ReDim oldHeight(1 To edgeCount)
    ReDim deleted(1 To edgeCount)
    made = 0

    On Error GoTo BackupFailed
    For index = 1 To edgeCount
        If edge(index) Is Nothing Then GoTo BackupClean
        names(index) = edge(index).Name
        stack(index) = edge(index).ZOrderPosition
        oldLeft(index) = edge(index).Left
        oldTop(index) = edge(index).Top
        oldWidth(index) = edge(index).Width
        oldHeight(index) = edge(index).Height
        Set copyShape = PPTK_DuplicateShape(edge(index))
        If copyShape Is Nothing Then GoTo BackupClean
        Set backups(index) = copyShape
        made = index
        copyShape.Visible = msoFalse
        copyShape.Left = oldLeft(index)
        copyShape.Top = oldTop(index)
    Next index

    On Error GoTo DeleteFailed
    For index = 1 To edgeCount
        edge(index).Delete
        deleted(index) = True
    Next index

    PPTK_DropShapes backups, made
    PPTK_DeleteEdgeSafely = True
    Exit Function

BackupFailed:
    Resume BackupClean
BackupClean:
    PPTK_DropShapes backups, made
    failReason = "A recovery copy could not be made, so nothing was deleted."
    Exit Function

DeleteFailed:
    Resume DeleteClean
DeleteClean:
    For index = 1 To edgeCount
        If deleted(index) Then
            PPTK_ReviveBackup backups(index), names(index), stack(index), oldLeft(index), oldTop(index), oldWidth(index), oldHeight(index)
        Else
            PPTK_DropShape backups(index)
        End If
    Next index
    failReason = "The edge could not be deleted. Every shape was kept."
End Function

Private Sub PPTK_ReviveBackup(ByVal backup As Shape, ByVal shapeName As String, ByVal stackPos As Long, ByVal leftPos As Single, ByVal topPos As Single, ByVal widthPts As Single, ByVal heightPts As Single)
    On Error Resume Next
    If backup Is Nothing Then Exit Sub
    backup.Visible = msoTrue
    backup.Width = widthPts
    backup.Height = heightPts
    backup.Left = leftPos
    backup.Top = topPos
    backup.Name = shapeName
    PPTK_SendToStack backup, stackPos
    Err.Clear
End Sub

Private Sub PPTK_SendToStack(ByVal shp As Shape, ByVal stackPos As Long)
    Dim guard As Long
    On Error Resume Next
    guard = 0
    Do While shp.ZOrderPosition > stackPos And guard < 1000
        shp.ZOrder msoSendBackward
        guard = guard + 1
        If Err.Number <> 0 Then Exit Do
    Loop
    Err.Clear
End Sub

Private Sub PPTK_DropShapes(ByRef pile() As Shape, ByVal made As Long)
    Dim index As Long
    For index = 1 To made
        PPTK_DropShape pile(index)
    Next index
End Sub

Private Sub PPTK_DropShape(ByVal shp As Shape)
    On Error Resume Next
    If Not shp Is Nothing Then shp.Delete
    Err.Clear
End Sub

' A hidden copy of a table, made before anything changes. If the table cannot
' keep its size after an edge is gone, the copy replaces it.
Private Function PPTK_HiddenCopy(ByVal shp As Shape, ByVal leftPos As Single, ByVal topPos As Single) As Shape
    Dim copyShape As Shape
    On Error GoTo Clean
    Set copyShape = PPTK_DuplicateShape(shp)
    If copyShape Is Nothing Then Exit Function
    copyShape.Visible = msoFalse
    copyShape.Left = leftPos
    copyShape.Top = topPos
    Set PPTK_HiddenCopy = copyShape
    Exit Function
Clean:
    Err.Clear
    PPTK_DropShape copyShape
    Set PPTK_HiddenCopy = Nothing
End Function

Private Function PPTK_SwapInCopy(ByVal shp As Shape, ByVal copyShape As Shape, ByVal leftPos As Single, ByVal topPos As Single) As Boolean
    Dim shapeName As String
    Dim stackPos As Long
    Dim originalGone As Boolean
    PPTK_SwapInCopy = False
    If copyShape Is Nothing Then Exit Function
    On Error GoTo Failed
    shapeName = shp.Name
    stackPos = shp.ZOrderPosition
    copyShape.Visible = msoTrue
    copyShape.Left = leftPos
    copyShape.Top = topPos
    shp.Delete
    originalGone = True
    On Error Resume Next
    copyShape.Name = shapeName
    PPTK_SendToStack copyShape, stackPos
    copyShape.Select
    Err.Clear
    PPTK_SwapInCopy = True
    Exit Function
Failed:
    Resume Cleanup
Cleanup:
    On Error Resume Next
    If Not originalGone Then copyShape.Visible = msoFalse
    Err.Clear
End Function

Private Sub PPTK_KeepAddShapes(ByVal bottomRow As Boolean)
    Dim items() As Shape
    Dim count As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim rowOf() As Long
    Dim colOf() As Long
    Dim lefts() As Single
    Dim tops() As Single
    Dim widths() As Single
    Dim heights() As Single
    Dim axisFixed() As Boolean
    Dim newLeft() As Single
    Dim newTop() As Single
    Dim newWidth() As Single
    Dim newHeight() As Single
    Dim copyLeft() As Single
    Dim copyTop() As Single
    Dim copyWidth() As Single
    Dim copyHeight() As Single
    Dim copies() As Shape
    Dim cellIndex() As Long
    Dim source As Shape
    Dim srcIndex As Long
    Dim index As Long
    Dim edgeCount As Long
    Dim made As Long
    Dim reason As String

    If Not PPTK_Recognize(items, count, rowCount, colCount, rowOf, colOf) Then
        PPTK_Tell PPTK_NotGridMessage()
        Exit Sub
    End If
    If Not PPTK_CaptureGrid(items, count, bottomRow, lefts, tops, widths, heights, axisFixed) Then
        PPTK_Tell "The grid could not be measured. The layout was not changed."
        Exit Sub
    End If
    If Not PPTK_PlanKeep(bottomRow, True, count, rowCount, colCount, rowOf, colOf, lefts, tops, widths, heights, axisFixed, newLeft, newTop, newWidth, newHeight, copyLeft, copyTop, copyWidth, copyHeight, reason) Then
        PPTK_Tell reason
        Exit Sub
    End If

    ReDim cellIndex(1 To rowCount, 1 To colCount)
    For index = 1 To count
        cellIndex(rowOf(index), colOf(index)) = index
    Next index
    If bottomRow Then edgeCount = colCount Else edgeCount = rowCount
    ReDim copies(1 To edgeCount)
    made = 0
    On Error GoTo Rollback
    For index = 1 To edgeCount
        If bottomRow Then
            srcIndex = cellIndex(rowCount, index)
        Else
            srcIndex = cellIndex(index, colCount)
        End If
        If srcIndex < 1 Or srcIndex > count Then GoTo Rollback
        Set source = items(srcIndex)
        If source Is Nothing Then GoTo Rollback
        On Error Resume Next
        Set copies(index) = PPTK_NewCell(source)
        If Err.Number <> 0 Then Set copies(index) = Nothing
        Err.Clear
        On Error GoTo Rollback
        If copies(index) Is Nothing Then GoTo Rollback
        made = made + 1
    Next index

    For index = 1 To count
        If Not PPTK_CommitShape(items(index), bottomRow, axisFixed(index), newLeft(index), newTop(index), newWidth(index), newHeight(index)) Then GoTo Rollback
    Next index
    For index = 1 To edgeCount
        If bottomRow Then
            srcIndex = cellIndex(rowCount, index)
        Else
            srcIndex = cellIndex(index, colCount)
        End If
        If srcIndex < 1 Or srcIndex > count Then GoTo Rollback
        If Not PPTK_CommitShape(copies(index), bottomRow, axisFixed(srcIndex), copyLeft(index), copyTop(index), copyWidth(index), copyHeight(index)) Then GoTo Rollback
    Next index

    On Error Resume Next
    items(1).Select
    For index = 2 To count
        items(index).Select False
    Next index
    For index = 1 To edgeCount
        copies(index).Select False
    Next index
    Exit Sub
Rollback:
    PPTK_RestoreShapes items, count, lefts, tops, widths, heights
    PPTK_DeleteCopies copies, made
    PPTK_Tell "The grid could not keep its original bounds. The layout was restored."
End Sub

Private Sub PPTK_KeepDeleteShapes(ByVal bottomRow As Boolean)
    Dim items() As Shape
    Dim count As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim rowOf() As Long
    Dim colOf() As Long
    Dim lefts() As Single
    Dim tops() As Single
    Dim widths() As Single
    Dim heights() As Single
    Dim axisFixed() As Boolean
    Dim newLeft() As Single
    Dim newTop() As Single
    Dim newWidth() As Single
    Dim newHeight() As Single
    Dim copyLeft() As Single
    Dim copyTop() As Single
    Dim copyWidth() As Single
    Dim copyHeight() As Single
    Dim index As Long
    Dim reason As String
    Dim edge() As Shape
    Dim edgeCount As Long

    If Not PPTK_Recognize(items, count, rowCount, colCount, rowOf, colOf) Then
        PPTK_Tell PPTK_NotGridMessage()
        Exit Sub
    End If
    If Not PPTK_LastEdgeGuard(bottomRow, rowCount, colCount) Then Exit Sub
    If Not PPTK_CaptureGrid(items, count, bottomRow, lefts, tops, widths, heights, axisFixed) Then
        PPTK_Tell "The grid could not be measured. The layout was not changed."
        Exit Sub
    End If
    If Not PPTK_PlanKeep(bottomRow, False, count, rowCount, colCount, rowOf, colOf, lefts, tops, widths, heights, axisFixed, newLeft, newTop, newWidth, newHeight, copyLeft, copyTop, copyWidth, copyHeight, reason) Then
        PPTK_Tell reason
        Exit Sub
    End If

    ' Survivors move first. That is reversible from the captured geometry.
    ' The edge is deleted last, through recovery copies, so a failure at any
    ' step leaves every shape on the slide.
    reason = "The grid could not keep its original bounds. The layout was restored."
    On Error GoTo Restore
    For index = 1 To count
        If bottomRow Then
            If rowOf(index) <> rowCount Then
                If Not PPTK_CommitShape(items(index), bottomRow, axisFixed(index), newLeft(index), newTop(index), newWidth(index), newHeight(index)) Then GoTo RestoreClean
            End If
        ElseIf colOf(index) <> colCount Then
            If Not PPTK_CommitShape(items(index), bottomRow, axisFixed(index), newLeft(index), newTop(index), newWidth(index), newHeight(index)) Then GoTo RestoreClean
        End If
    Next index
    PPTK_CollectEdge items, count, rowOf, colOf, rowCount, colCount, bottomRow, edge, edgeCount
    If Not PPTK_DeleteEdgeSafely(edge, edgeCount, reason) Then GoTo RestoreClean
    PPTK_SelectSurvivors items, count, rowOf, colOf, rowCount, colCount, bottomRow
    Exit Sub
Restore:
    Resume RestoreClean
RestoreClean:
    PPTK_RestoreShapes items, count, lefts, tops, widths, heights
    PPTK_Tell reason
End Sub

Private Sub PPTK_KeepAddTable(ByVal shp As Shape, ByVal bottomRow As Boolean)
    Dim oldWidths() As Single
    Dim oldHeights() As Single
    Dim newSizes() As Single
    Dim colCount As Long
    Dim rowCount As Long
    Dim savedLeft As Single
    Dim savedTop As Single
    Dim savedWidth As Single
    Dim savedHeight As Single
    Dim index As Long
    Dim reason As String
    Dim added As Boolean
    Dim target As Single

    If PPTK_TableMerged(shp) Then
        PPTK_Tell "This table has merged cells. The layout was not changed."
        Exit Sub
    End If

    On Error GoTo MeasureFailed
    savedLeft = shp.Left
    savedTop = shp.Top
    savedWidth = shp.Width
    savedHeight = shp.Height
    colCount = shp.Table.Columns.Count
    rowCount = shp.Table.Rows.Count
    ReDim oldWidths(1 To colCount)
    ReDim oldHeights(1 To rowCount)
    For index = 1 To colCount
        oldWidths(index) = shp.Table.Columns(index).Width
    Next index
    For index = 1 To rowCount
        oldHeights(index) = shp.Table.Rows(index).Height
    Next index
    On Error GoTo 0

    If bottomRow Then
        target = savedHeight
        ReDim newSizes(1 To rowCount + 1)
        For index = 1 To rowCount
            newSizes(index) = oldHeights(index)
        Next index
        newSizes(rowCount + 1) = oldHeights(rowCount)
        If Not PPTK_FitSizes(newSizes, rowCount + 1, target, True, reason) Then
            PPTK_Tell reason
            Exit Sub
        End If
    Else
        target = savedWidth
        ReDim newSizes(1 To colCount + 1)
        For index = 1 To colCount
            newSizes(index) = oldWidths(index)
        Next index
        newSizes(colCount + 1) = oldWidths(colCount)
        If Not PPTK_FitSizes(newSizes, colCount + 1, target, False, reason) Then
            PPTK_Tell reason
            Exit Sub
        End If
    End If

    On Error GoTo Failed
    If bottomRow Then
        shp.Table.Rows.Add
        added = True
        If shp.Table.Rows.Count <> rowCount + 1 Then GoTo FailedClean
        If Not PPTK_WriteTableAxis(shp, newSizes, rowCount + 1, savedHeight, True, savedLeft, savedTop) Then GoTo FailedClean
        If Abs(shp.Width - savedWidth) > BoundSlack Then
            PPTK_RestoreTableAxis shp, oldWidths, colCount, False, savedLeft, savedTop
            If Not PPTK_WriteTableAxis(shp, newSizes, rowCount + 1, savedHeight, True, savedLeft, savedTop) Then GoTo FailedClean
        End If
        If Abs(shp.Height - savedHeight) > BoundSlack Or Abs(shp.Width - savedWidth) > BoundSlack Then GoTo FailedClean
    Else
        shp.Table.Columns.Add
        added = True
        If shp.Table.Columns.Count <> colCount + 1 Then GoTo FailedClean
        If Not PPTK_WriteTableAxis(shp, newSizes, colCount + 1, savedWidth, False, savedLeft, savedTop) Then GoTo FailedClean
        If Abs(shp.Height - savedHeight) > BoundSlack Then
            PPTK_RestoreTableAxis shp, oldHeights, rowCount, True, savedLeft, savedTop
            If Not PPTK_WriteTableAxis(shp, newSizes, colCount + 1, savedWidth, False, savedLeft, savedTop) Then GoTo FailedClean
        End If
        If Abs(shp.Width - savedWidth) > BoundSlack Or Abs(shp.Height - savedHeight) > BoundSlack Then GoTo FailedClean
    End If
    shp.Left = savedLeft
    shp.Top = savedTop
    Exit Sub
MeasureFailed:
    PPTK_Tell "The table could not be measured. The layout was not changed."
    Exit Sub
Failed:
    Resume FailedClean
FailedClean:
    On Error Resume Next
    If added Then
        If bottomRow Then
            If shp.Table.Rows.Count = rowCount + 1 Then shp.Table.Rows(shp.Table.Rows.Count).Delete
        Else
            If shp.Table.Columns.Count = colCount + 1 Then shp.Table.Columns(shp.Table.Columns.Count).Delete
        End If
    End If
    PPTK_RestoreTableAxis shp, oldWidths, colCount, False, savedLeft, savedTop
    PPTK_RestoreTableAxis shp, oldHeights, rowCount, True, savedLeft, savedTop
    shp.Left = savedLeft
    shp.Top = savedTop
    PPTK_Tell "The table could not keep its original size. The change was undone."
End Sub

Private Sub PPTK_KeepDeleteTable(ByVal shp As Shape, ByVal bottomRow As Boolean)
    Dim oldWidths() As Single
    Dim oldHeights() As Single
    Dim newSizes() As Single
    Dim colCount As Long
    Dim rowCount As Long
    Dim savedLeft As Single
    Dim savedTop As Single
    Dim savedWidth As Single
    Dim savedHeight As Single
    Dim index As Long
    Dim reason As String
    Dim sum As Single
    Dim removed As Boolean
    Dim before As Long

    Dim backup As Shape

    On Error GoTo MeasureFailed
    If bottomRow Then
        If shp.Table.Rows.Count <= 1 Then
            PPTK_Tell "The last row of a table is not deleted."
            Exit Sub
        End If
    Else
        If shp.Table.Columns.Count <= 1 Then
            PPTK_Tell "The last column of a table is not deleted."
            Exit Sub
        End If
    End If
    On Error GoTo 0

    If PPTK_TableMerged(shp) Then
        PPTK_Tell "This table has merged cells. The layout was not changed."
        Exit Sub
    End If

    On Error GoTo MeasureFailed
    savedLeft = shp.Left
    savedTop = shp.Top
    savedWidth = shp.Width
    savedHeight = shp.Height
    colCount = shp.Table.Columns.Count
    rowCount = shp.Table.Rows.Count
    If bottomRow Then
        If rowCount <= 1 Then Exit Sub
    Else
        If colCount <= 1 Then Exit Sub
    End If
    ReDim oldWidths(1 To colCount)
    ReDim oldHeights(1 To rowCount)
    For index = 1 To colCount
        oldWidths(index) = shp.Table.Columns(index).Width
    Next index
    For index = 1 To rowCount
        oldHeights(index) = shp.Table.Rows(index).Height
    Next index
    On Error GoTo 0

    If bottomRow Then
        ReDim newSizes(1 To rowCount - 1)
        For index = 1 To rowCount - 1
            newSizes(index) = oldHeights(index)
        Next index
        If Not PPTK_FitSizes(newSizes, rowCount - 1, savedHeight, True, reason) Then
            PPTK_Tell reason
            Exit Sub
        End If
    Else
        ReDim newSizes(1 To colCount - 1)
        For index = 1 To colCount - 1
            newSizes(index) = oldWidths(index)
        Next index
        If Not PPTK_FitSizes(newSizes, colCount - 1, savedWidth, False, reason) Then
            PPTK_Tell reason
            Exit Sub
        End If
    End If

    ' A hidden copy is made before anything changes. Cell content in the
    ' deleted edge cannot be rebuilt from Rows.Add or Columns.Add, so the
    ' copy is the only way back once the edge is gone.
    Set backup = PPTK_HiddenCopy(shp, savedLeft, savedTop)
    If backup Is Nothing Then
        PPTK_Tell "A recovery copy of the table could not be made, so nothing was deleted."
        Exit Sub
    End If

    On Error GoTo Failed
    If bottomRow Then
        For index = 1 To rowCount - 1
            shp.Table.Rows(index).Height = newSizes(index)
        Next index
        sum = 0
        For index = 1 To rowCount - 1
            sum = sum + shp.Table.Rows(index).Height
        Next index
        If Abs(sum - savedHeight) > BoundSlack Then GoTo FailedClean
        before = shp.Table.Rows.Count
        On Error Resume Next
        shp.Table.Rows(rowCount).Delete
        If shp.Table.Rows.Count < before Then removed = True
        Err.Clear
        On Error GoTo Failed
        If Not removed Then GoTo FailedClean
        shp.Left = savedLeft
        shp.Top = savedTop
        If Abs(shp.Height - savedHeight) > BoundSlack Then
            shp.Table.Rows(rowCount - 1).Height = shp.Table.Rows(rowCount - 1).Height + (savedHeight - shp.Height)
            shp.Left = savedLeft
            shp.Top = savedTop
        End If
        If Abs(shp.Height - savedHeight) > BoundSlack Then GoTo FailedClean
        If Abs(shp.Width - savedWidth) > BoundSlack Then
            PPTK_RestoreTableAxis shp, oldWidths, colCount, False, savedLeft, savedTop
            shp.Left = savedLeft
            shp.Top = savedTop
        End If
        If Abs(shp.Height - savedHeight) > BoundSlack Or Abs(shp.Width - savedWidth) > BoundSlack Then GoTo FailedClean
    Else
        For index = 1 To colCount - 1
            shp.Table.Columns(index).Width = newSizes(index)
        Next index
        sum = 0
        For index = 1 To colCount - 1
            sum = sum + shp.Table.Columns(index).Width
        Next index
        If Abs(sum - savedWidth) > BoundSlack Then GoTo FailedClean
        before = shp.Table.Columns.Count
        On Error Resume Next
        shp.Table.Columns(colCount).Delete
        If shp.Table.Columns.Count < before Then removed = True
        Err.Clear
        On Error GoTo Failed
        If Not removed Then GoTo FailedClean
        shp.Left = savedLeft
        shp.Top = savedTop
        If Abs(shp.Width - savedWidth) > BoundSlack Then
            shp.Table.Columns(colCount - 1).Width = shp.Table.Columns(colCount - 1).Width + (savedWidth - shp.Width)
            shp.Left = savedLeft
            shp.Top = savedTop
        End If
        If Abs(shp.Width - savedWidth) > BoundSlack Then GoTo FailedClean
        If Abs(shp.Height - savedHeight) > BoundSlack Then
            PPTK_RestoreTableAxis shp, oldHeights, rowCount, True, savedLeft, savedTop
            shp.Left = savedLeft
            shp.Top = savedTop
        End If
        If Abs(shp.Width - savedWidth) > BoundSlack Or Abs(shp.Height - savedHeight) > BoundSlack Then GoTo FailedClean
    End If
    shp.Left = savedLeft
    shp.Top = savedTop
    PPTK_DropShape backup
    Exit Sub
MeasureFailed:
    PPTK_Tell "The table could not be measured. The layout was not changed."
    Exit Sub
Failed:
    Resume FailedClean
FailedClean:
    On Error Resume Next
    If removed Then
        ' The edge is gone. Put the untouched copy back in place of the table.
        If PPTK_SwapInCopy(shp, backup, savedLeft, savedTop) Then
            PPTK_Tell "The table could not keep its original size. The original table was restored."
        Else
            PPTK_Tell "The table could not keep its original size, and it could not be swapped back. " & _
                "A hidden copy of the original table is on this slide (Selection Pane)."
        End If
        Exit Sub
    End If
    If shp.Table.Columns.Count = colCount Then PPTK_RestoreTableAxis shp, oldWidths, colCount, False, savedLeft, savedTop
    If shp.Table.Rows.Count = rowCount Then PPTK_RestoreTableAxis shp, oldHeights, rowCount, True, savedLeft, savedTop
    shp.Left = savedLeft
    shp.Top = savedTop
    PPTK_DropShape backup
    PPTK_Tell "The table could not keep its original size. The change was undone."
End Sub

Private Function PPTK_CaptureGrid(ByRef items() As Shape, ByVal count As Long, ByVal bottomRow As Boolean, ByRef lefts() As Single, ByRef tops() As Single, ByRef widths() As Single, ByRef heights() As Single, ByRef axisFixed() As Boolean) As Boolean
    Dim index As Long
    ReDim lefts(1 To count)
    ReDim tops(1 To count)
    ReDim widths(1 To count)
    ReDim heights(1 To count)
    ReDim axisFixed(1 To count)
    On Error GoTo Clean
    For index = 1 To count
        lefts(index) = items(index).Left
        tops(index) = items(index).Top
        widths(index) = items(index).Width
        heights(index) = items(index).Height
        axisFixed(index) = PPTK_AxisFixed(items(index), bottomRow)
    Next index
    PPTK_CaptureGrid = True
    Exit Function
Clean:
    PPTK_CaptureGrid = False
    Err.Clear
End Function

Private Function PPTK_PlanKeep(ByVal bottomRow As Boolean, ByVal adding As Boolean, ByVal count As Long, ByVal rowCount As Long, ByVal colCount As Long, ByRef rowOf() As Long, ByRef colOf() As Long, ByRef lefts() As Single, ByRef tops() As Single, ByRef widths() As Single, ByRef heights() As Single, ByRef axisFixed() As Boolean, ByRef newLeft() As Single, ByRef newTop() As Single, ByRef newWidth() As Single, ByRef newHeight() As Single, ByRef copyLeft() As Single, ByRef copyTop() As Single, ByRef copyWidth() As Single, ByRef copyHeight() As Single, ByRef failReason As String) As Boolean
    Dim cellIndex() As Long
    Dim slotPos() As Single
    Dim slotSize() As Single
    Dim slotSeen() As Boolean
    Dim fixedFloor() As Single
    Dim flexPart() As Single
    Dim gap() As Single
    Dim extFix() As Single
    Dim extFlex() As Single
    Dim extGap() As Single
    Dim outSlot() As Single
    Dim outGap() As Single
    Dim outPos() As Single
    Dim index As Long
    Dim row As Long
    Dim col As Long
    Dim axis As Long
    Dim axisCount As Long
    Dim newCount As Long
    Dim filled As Long
    Dim pos As Single
    Dim endPos As Single
    Dim size As Single
    Dim boundMin As Single
    Dim boundMax As Single
    Dim target As Single
    Dim newGap As Single
    Dim flexSum As Single
    Dim fixSum As Single
    Dim factor As Single
    Dim leftover As Single
    Dim share As Single
    Dim sumGaps As Single
    Dim used As Single
    Dim remainder As Single
    Dim cursor As Single
    Dim src As Long
    Dim prevEnd As Single
    Dim lastBand As Long
    Dim keepLimit As Long

    failReason = "The grid cannot keep its original bounds."
    PPTK_PlanKeep = False
    If count < 1 Or rowCount < 1 Or colCount < 1 Then Exit Function
    If Not adding Then
        If bottomRow And rowCount <= 1 Then Exit Function
        If Not bottomRow And colCount <= 1 Then Exit Function
    End If

    ReDim cellIndex(1 To rowCount, 1 To colCount)
    For index = 1 To count
        cellIndex(rowOf(index), colOf(index)) = index
    Next index
    filled = 0
    For row = 1 To rowCount
        For col = 1 To colCount
            If cellIndex(row, col) <> 0 Then filled = filled + 1
        Next col
    Next row
    If filled <> count Then Exit Function

    If bottomRow Then axisCount = rowCount Else axisCount = colCount
    ReDim slotPos(1 To axisCount)
    ReDim slotSize(1 To axisCount)
    ReDim slotSeen(1 To axisCount)
    ReDim fixedFloor(1 To axisCount)
    ReDim flexPart(1 To axisCount)
    For index = 1 To count
        If bottomRow Then
            axis = rowOf(index)
            pos = tops(index)
            endPos = tops(index) + heights(index)
            size = heights(index)
        Else
            axis = colOf(index)
            pos = lefts(index)
            endPos = lefts(index) + widths(index)
            size = widths(index)
        End If
        If Not slotSeen(axis) Then
            slotSeen(axis) = True
            slotPos(axis) = pos
            slotSize(axis) = endPos
        Else
            If pos < slotPos(axis) Then slotPos(axis) = pos
            If endPos > slotSize(axis) Then slotSize(axis) = endPos
        End If
        If axisFixed(index) Then
            If size > fixedFloor(axis) Then fixedFloor(axis) = size
        End If
    Next index
    For axis = 1 To axisCount
        If Not slotSeen(axis) Then Exit Function
        slotSize(axis) = slotSize(axis) - slotPos(axis)
        If slotSize(axis) < 0 Then slotSize(axis) = 0
        flexPart(axis) = slotSize(axis) - fixedFloor(axis)
        If flexPart(axis) < 0 Then flexPart(axis) = 0
    Next axis

    If bottomRow Then
        boundMin = tops(1)
        boundMax = tops(1) + heights(1)
        For index = 2 To count
            If tops(index) < boundMin Then boundMin = tops(index)
            If tops(index) + heights(index) > boundMax Then boundMax = tops(index) + heights(index)
        Next index
    Else
        boundMin = lefts(1)
        boundMax = lefts(1) + widths(1)
        For index = 2 To count
            If lefts(index) < boundMin Then boundMin = lefts(index)
            If lefts(index) + widths(index) > boundMax Then boundMax = lefts(index) + widths(index)
        Next index
    End If
    target = boundMax - boundMin
    If target <= 0.05 Then Exit Function

    newGap = 0
    If axisCount >= 2 Then
        ReDim gap(1 To axisCount - 1)
        For axis = 1 To axisCount - 1
            gap(axis) = slotPos(axis + 1) - (slotPos(axis) + slotSize(axis))
            If gap(axis) < 0 Then gap(axis) = 0
        Next axis
        newGap = gap(axisCount - 1)
    End If

    If adding Then
        newCount = axisCount + 1
    Else
        newCount = axisCount - 1
    End If
    If newCount < 1 Then Exit Function
    ReDim extFix(1 To newCount)
    ReDim extFlex(1 To newCount)
    keepLimit = newCount
    If adding Then keepLimit = axisCount
    For axis = 1 To keepLimit
        extFix(axis) = fixedFloor(axis)
        extFlex(axis) = flexPart(axis)
    Next axis
    If adding Then
        extFix(newCount) = fixedFloor(axisCount)
        extFlex(newCount) = flexPart(axisCount)
    End If
    If newCount >= 2 Then
        ReDim extGap(1 To newCount - 1)
        If adding Then
            For axis = 1 To axisCount - 1
                extGap(axis) = gap(axis)
            Next axis
            extGap(newCount - 1) = newGap
        Else
            For axis = 1 To newCount - 1
                extGap(axis) = gap(axis)
            Next axis
        End If
    End If

    flexSum = 0
    fixSum = 0
    For axis = 1 To newCount
        fixSum = fixSum + extFix(axis)
        flexSum = flexSum + extFlex(axis)
    Next axis
    If newCount >= 2 Then
        For axis = 1 To newCount - 1
            flexSum = flexSum + extGap(axis)
        Next axis
    End If
    If fixSum > target + BoundSlack Then
        failReason = PPTK_FixedFitMessage(bottomRow)
        Exit Function
    End If

    ReDim outSlot(1 To newCount)
    If newCount >= 2 Then ReDim outGap(1 To newCount - 1)
    If flexSum <= 0.05 Then
        leftover = target - fixSum
        If leftover < 0 Then leftover = 0
        If newCount = 1 Then
            outSlot(1) = target
        Else
            For axis = 1 To newCount
                outSlot(axis) = extFix(axis)
            Next axis
            share = leftover / (newCount - 1)
            sumGaps = 0
            For axis = 1 To newCount - 2
                outGap(axis) = share
                sumGaps = sumGaps + share
            Next axis
            outGap(newCount - 1) = leftover - sumGaps
            If outGap(newCount - 1) < 0 Then outGap(newCount - 1) = 0
        End If
    Else
        factor = (target - fixSum) / flexSum
        If factor < 0 Then factor = 0
        For axis = 1 To newCount
            outSlot(axis) = extFix(axis) + extFlex(axis) * factor
        Next axis
        If newCount >= 2 Then
            For axis = 1 To newCount - 1
                outGap(axis) = extGap(axis) * factor
                If outGap(axis) < 0 Then outGap(axis) = 0
            Next axis
        End If
        used = 0
        For axis = 1 To newCount
            used = used + outSlot(axis)
        Next axis
        If newCount >= 2 Then
            For axis = 1 To newCount - 1
                used = used + outGap(axis)
            Next axis
        End If
        remainder = target - used
        If outSlot(newCount) + remainder >= extFix(newCount) - 0.05 Then
            outSlot(newCount) = outSlot(newCount) + remainder
        ElseIf newCount >= 2 Then
            outGap(newCount - 1) = outGap(newCount - 1) + remainder
            If outGap(newCount - 1) < -0.05 Then Exit Function
            If outGap(newCount - 1) < 0 Then outGap(newCount - 1) = 0
        Else
            Exit Function
        End If
    End If

    For axis = 1 To newCount
        If outSlot(axis) < extFix(axis) - 0.05 Or outSlot(axis) < 0 Then Exit Function
    Next axis
    If newCount >= 2 Then
        For axis = 1 To newCount - 1
            If outGap(axis) < -0.05 Then Exit Function
        Next axis
    End If

    ReDim outPos(1 To newCount)
    cursor = boundMin
    For axis = 1 To newCount
        outPos(axis) = cursor
        cursor = cursor + outSlot(axis)
        If axis < newCount Then cursor = cursor + outGap(axis)
    Next axis
    If Abs(cursor - boundMax) > BoundSlack Then
        outSlot(newCount) = outSlot(newCount) + (boundMax - cursor)
        If outSlot(newCount) < extFix(newCount) - 0.05 Then Exit Function
        cursor = boundMin
        For axis = 1 To newCount
            outPos(axis) = cursor
            cursor = cursor + outSlot(axis)
            If axis < newCount Then cursor = cursor + outGap(axis)
        Next axis
    End If
    If Abs(cursor - boundMax) > BoundSlack Then Exit Function

    ReDim newLeft(1 To count)
    ReDim newTop(1 To count)
    ReDim newWidth(1 To count)
    ReDim newHeight(1 To count)
    For index = 1 To count
        newLeft(index) = lefts(index)
        newTop(index) = tops(index)
        newWidth(index) = widths(index)
        newHeight(index) = heights(index)
        If bottomRow Then
            axis = rowOf(index)
            If Not adding And axis = rowCount Then GoTo NextPlace
            PPTK_PlaceOnAxis tops(index), heights(index), slotPos(axis), slotSize(axis), outPos(axis), outSlot(axis), axisFixed(index), newTop(index), newHeight(index)
            If Not PPTK_AxisSizeOk(heights(index), newHeight(index), outSlot(axis), axisFixed(index), failReason, bottomRow) Then Exit Function
        Else
            axis = colOf(index)
            If Not adding And axis = colCount Then GoTo NextPlace
            PPTK_PlaceOnAxis lefts(index), widths(index), slotPos(axis), slotSize(axis), outPos(axis), outSlot(axis), axisFixed(index), newLeft(index), newWidth(index)
            If Not PPTK_AxisSizeOk(widths(index), newWidth(index), outSlot(axis), axisFixed(index), failReason, bottomRow) Then Exit Function
        End If
NextPlace:
    Next index

    If bottomRow Then
        ReDim copyLeft(1 To colCount)
        ReDim copyTop(1 To colCount)
        ReDim copyWidth(1 To colCount)
        ReDim copyHeight(1 To colCount)
    Else
        ReDim copyLeft(1 To rowCount)
        ReDim copyTop(1 To rowCount)
        ReDim copyWidth(1 To rowCount)
        ReDim copyHeight(1 To rowCount)
    End If
    If adding Then
        If bottomRow Then
            For col = 1 To colCount
                src = cellIndex(rowCount, col)
                If src = 0 Then Exit Function
                copyLeft(col) = lefts(src)
                copyWidth(col) = widths(src)
                PPTK_PlaceOnAxis tops(src), heights(src), slotPos(rowCount), slotSize(rowCount), outPos(newCount), outSlot(newCount), axisFixed(src), copyTop(col), copyHeight(col)
                If Not PPTK_AxisSizeOk(heights(src), copyHeight(col), outSlot(newCount), axisFixed(src), failReason, bottomRow) Then Exit Function
            Next col
        Else
            For row = 1 To rowCount
                src = cellIndex(row, colCount)
                If src = 0 Then Exit Function
                copyTop(row) = tops(src)
                copyHeight(row) = heights(src)
                PPTK_PlaceOnAxis lefts(src), widths(src), slotPos(colCount), slotSize(colCount), outPos(newCount), outSlot(newCount), axisFixed(src), copyLeft(row), copyWidth(row)
                If Not PPTK_AxisSizeOk(widths(src), copyWidth(row), outSlot(newCount), axisFixed(src), failReason, bottomRow) Then Exit Function
            Next row
        End If
    End If

    If bottomRow Then
        lastBand = rowCount - 1
        If adding Then lastBand = rowCount + 1
        For col = 1 To colCount
            prevEnd = -10000000#
            For row = 1 To lastBand
                If adding And row = rowCount + 1 Then
                    pos = copyTop(col)
                    size = copyHeight(col)
                Else
                    src = cellIndex(row, col)
                    pos = newTop(src)
                    size = newHeight(src)
                End If
                If pos < prevEnd - 0.5 Then
                    failReason = "The grid would overlap, so it was not changed."
                    Exit Function
                End If
                If pos < boundMin - BoundSlack Or pos + size > boundMax + BoundSlack Then Exit Function
                prevEnd = pos + size
            Next row
        Next col
    Else
        lastBand = colCount - 1
        If adding Then lastBand = colCount + 1
        For row = 1 To rowCount
            prevEnd = -10000000#
            For col = 1 To lastBand
                If adding And col = colCount + 1 Then
                    pos = copyLeft(row)
                    size = copyWidth(row)
                Else
                    src = cellIndex(row, col)
                    pos = newLeft(src)
                    size = newWidth(src)
                End If
                If pos < prevEnd - 0.5 Then
                    failReason = "The grid would overlap, so it was not changed."
                    Exit Function
                End If
                If pos < boundMin - BoundSlack Or pos + size > boundMax + BoundSlack Then Exit Function
                prevEnd = pos + size
            Next col
        Next row
    End If

    PPTK_PlanKeep = True
End Function

Private Sub PPTK_PlaceOnAxis(ByVal oldPos As Single, ByVal oldSize As Single, ByVal oldSlotPos As Single, ByVal oldSlotSize As Single, ByVal newSlotPos As Single, ByVal newSlotSize As Single, ByVal isFixed As Boolean, ByRef newPos As Single, ByRef newSize As Single)
    Dim centerFrac As Single
    Dim scaleSlot As Single
    If isFixed Then
        newSize = oldSize
        If oldSlotSize <= 0.05 Then
            newPos = newSlotPos + (newSlotSize - oldSize) / 2
        Else
            centerFrac = ((oldPos - oldSlotPos) + oldSize / 2) / oldSlotSize
            newPos = newSlotPos + centerFrac * newSlotSize - oldSize / 2
        End If
        If newPos < newSlotPos Then newPos = newSlotPos
        If newPos + newSize > newSlotPos + newSlotSize Then newPos = newSlotPos + newSlotSize - newSize
        If newPos < newSlotPos Then newPos = newSlotPos
    ElseIf oldSlotSize <= 0.05 Then
        newSize = newSlotSize
        newPos = newSlotPos
    Else
        scaleSlot = newSlotSize / oldSlotSize
        newSize = oldSize * scaleSlot
        newPos = newSlotPos + (oldPos - oldSlotPos) * scaleSlot
    End If
End Sub

Private Function PPTK_AxisSizeOk(ByVal oldSize As Single, ByVal newSize As Single, ByVal slotSize As Single, ByVal isFixed As Boolean, ByRef failReason As String, ByVal bottomRow As Boolean) As Boolean
    PPTK_AxisSizeOk = False
    If newSize < -0.05 Then
        failReason = "The grid cannot keep its original bounds."
        Exit Function
    End If
    If isFixed Then
        If oldSize > slotSize + BoundSlack Then
            failReason = PPTK_FixedFitMessage(bottomRow)
            Exit Function
        End If
    ElseIf oldSize >= MinAxisSize And newSize < MinAxisSize Then
        failReason = "Shapes would become too small to keep the original bounds."
        Exit Function
    End If
    PPTK_AxisSizeOk = True
End Function

Private Function PPTK_FixedFitMessage(ByVal bottomRow As Boolean) As String
    If bottomRow Then
        PPTK_FixedFitMessage = "Fixed shapes cannot fit in the original height."
    Else
        PPTK_FixedFitMessage = "Fixed shapes cannot fit in the original width."
    End If
End Function

Private Function PPTK_CommitShape(ByVal shp As Shape, ByVal bottomRow As Boolean, ByVal isFixed As Boolean, ByVal newLeft As Single, ByVal newTop As Single, ByVal newWidth As Single, ByVal newHeight As Single) As Boolean
    On Error GoTo Fail
    If shp Is Nothing Then GoTo Fail
    If bottomRow Then
        If Not isFixed Then shp.Height = newHeight
        shp.Top = newTop
        shp.Left = newLeft
        If Not isFixed Then shp.Width = newWidth
    Else
        If Not isFixed Then shp.Width = newWidth
        shp.Left = newLeft
        shp.Top = newTop
        If Not isFixed Then shp.Height = newHeight
    End If
    If Abs(shp.Left - newLeft) > BoundSlack Then GoTo Fail
    If Abs(shp.Top - newTop) > BoundSlack Then GoTo Fail
    If Abs(shp.Width - newWidth) > BoundSlack Then GoTo Fail
    If Abs(shp.Height - newHeight) > BoundSlack Then GoTo Fail
    PPTK_CommitShape = True
    Exit Function
Fail:
    PPTK_CommitShape = False
    Err.Clear
End Function

Private Sub PPTK_RestoreShapes(ByRef items() As Shape, ByVal count As Long, ByRef lefts() As Single, ByRef tops() As Single, ByRef widths() As Single, ByRef heights() As Single)
    Dim index As Long
    On Error Resume Next
    For index = 1 To count
        If Not items(index) Is Nothing Then
            items(index).Width = widths(index)
            items(index).Height = heights(index)
            items(index).Left = lefts(index)
            items(index).Top = tops(index)
        End If
    Next index
    Err.Clear
End Sub

Private Sub PPTK_DeleteCopies(ByRef copies() As Shape, ByVal made As Long)
    Dim index As Long
    On Error Resume Next
    For index = 1 To made
        If Not copies(index) Is Nothing Then copies(index).Delete
    Next index
    Err.Clear
End Sub

Private Function PPTK_FitSizes(ByRef sizes() As Single, ByVal count As Long, ByVal target As Single, ByVal bottomRow As Boolean, ByRef failReason As String) As Boolean
    Dim index As Long
    Dim total As Single
    Dim factor As Single
    Dim sum As Single
    PPTK_FitSizes = False
    If bottomRow Then
        failReason = "The rows cannot fit in the original height."
    Else
        failReason = "The columns cannot fit in the original width."
    End If
    If count < 1 Or target <= 0.05 Then Exit Function
    total = 0
    For index = 1 To count
        If sizes(index) < 0 Then sizes(index) = 0
        total = total + sizes(index)
    Next index
    If total <= 0.05 Then
        For index = 1 To count
            sizes(index) = target / count
        Next index
    Else
        factor = target / total
        For index = 1 To count
            sizes(index) = sizes(index) * factor
        Next index
    End If
    If count = 1 Then
        sizes(1) = target
    Else
        sum = 0
        For index = 1 To count - 1
            sum = sum + sizes(index)
        Next index
        sizes(count) = target - sum
    End If
    For index = 1 To count
        If sizes(index) < MinAxisSize Then Exit Function
    Next index
    PPTK_FitSizes = True
End Function

Private Function PPTK_WriteTableAxis(ByVal shp As Shape, ByRef sizes() As Single, ByVal count As Long, ByVal target As Single, ByVal bottomRow As Boolean, ByVal savedLeft As Single, ByVal savedTop As Single) As Boolean
    Dim index As Long
    Dim others As Single
    Dim pass As Long
    Dim actual As Single
    On Error GoTo Clean
    If bottomRow Then
        If shp.Table.Rows.Count <> count Then Exit Function
    Else
        If shp.Table.Columns.Count <> count Then Exit Function
    End If
    If count = 1 Then
        If bottomRow Then
            shp.Table.Rows(1).Height = target
        Else
            shp.Table.Columns(1).Width = target
        End If
    Else
        For index = 1 To count - 1
            If bottomRow Then
                shp.Table.Rows(index).Height = sizes(index)
            Else
                shp.Table.Columns(index).Width = sizes(index)
            End If
        Next index
        others = 0
        For index = 1 To count - 1
            If bottomRow Then
                others = others + shp.Table.Rows(index).Height
            Else
                others = others + shp.Table.Columns(index).Width
            End If
        Next index
        If bottomRow Then
            shp.Table.Rows(count).Height = target - others
        Else
            shp.Table.Columns(count).Width = target - others
        End If
    End If
    For pass = 1 To 4
        shp.Left = savedLeft
        shp.Top = savedTop
        If bottomRow Then actual = shp.Height Else actual = shp.Width
        If Abs(actual - target) <= BoundSlack Then
            PPTK_WriteTableAxis = True
            Exit Function
        End If
        If bottomRow Then
            shp.Table.Rows(count).Height = shp.Table.Rows(count).Height + (target - actual)
        Else
            shp.Table.Columns(count).Width = shp.Table.Columns(count).Width + (target - actual)
        End If
    Next pass
    shp.Left = savedLeft
    shp.Top = savedTop
    If bottomRow Then actual = shp.Height Else actual = shp.Width
    PPTK_WriteTableAxis = Abs(actual - target) <= BoundSlack
    Exit Function
Clean:
    PPTK_WriteTableAxis = False
    Err.Clear
End Function

Private Sub PPTK_RestoreTableAxis(ByVal shp As Shape, ByRef sizes() As Single, ByVal count As Long, ByVal bottomRow As Boolean, ByVal savedLeft As Single, ByVal savedTop As Single)
    Dim index As Long
    On Error Resume Next
    If count < 1 Then Exit Sub
    If bottomRow Then
        If shp.Table.Rows.Count <> count Then Exit Sub
        For index = 1 To count
            shp.Table.Rows(index).Height = sizes(index)
        Next index
    Else
        If shp.Table.Columns.Count <> count Then Exit Sub
        For index = 1 To count
            shp.Table.Columns(index).Width = sizes(index)
        Next index
    End If
    shp.Left = savedLeft
    shp.Top = savedTop
    Err.Clear
End Sub

Private Function PPTK_TableMerged(ByVal shp As Shape) As Boolean
    Dim tbl As Table
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim cellBox As Shape
    ' PowerPoint's Cell has no RowSpan or ColumnSpan. A merged cell reports
    ' the merged area's size, so a cell wider than its column or taller than
    ' its row belongs to a merged area.
    On Error GoTo Merged
    If shp Is Nothing Then GoTo Merged
    If Not shp.HasTable Then GoTo Merged
    Set tbl = shp.Table
    For rowIndex = 1 To tbl.Rows.Count
        For colIndex = 1 To tbl.Columns.Count
            Set cellBox = tbl.Cell(rowIndex, colIndex).Shape
            If Abs(cellBox.Width - tbl.Columns(colIndex).Width) > 1 Then GoTo Merged
            If Abs(cellBox.Height - tbl.Rows(rowIndex).Height) > 1 Then GoTo Merged
        Next colIndex
    Next rowIndex
    PPTK_TableMerged = False
    Exit Function
Merged:
    PPTK_TableMerged = True
    Err.Clear
End Function

Private Function PPTK_AxisFixed(ByVal shp As Shape, ByVal bottomRow As Boolean) As Boolean
    Dim locked As MsoTriState
    PPTK_AxisFixed = False
    On Error Resume Next
    If shp Is Nothing Then
        PPTK_AxisFixed = True
        Err.Clear
        Exit Function
    End If
    If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
        PPTK_AxisFixed = True
        Err.Clear
        Exit Function
    End If
    If shp.Type = msoGroup Then
        Err.Clear
        locked = shp.LockAspectRatio
        If Err.Number = 0 Then
            If locked = msoTrue Then
                PPTK_AxisFixed = True
                Err.Clear
                Exit Function
            End If
        End If
        If PPTK_GroupAxisFixed(shp, bottomRow) Or PPTK_AutoSizeBlocks(shp, bottomRow) Then
            PPTK_AxisFixed = True
        Else
            PPTK_AxisFixed = False
        End If
        Err.Clear
        Exit Function
    End If
    Err.Clear
    locked = shp.LockAspectRatio
    If Err.Number = 0 Then
        If locked = msoTrue Then
            PPTK_AxisFixed = True
            Err.Clear
            Exit Function
        End If
    End If
    If PPTK_AutoSizeBlocks(shp, bottomRow) Then PPTK_AxisFixed = True
    Err.Clear
End Function

Private Function PPTK_GroupAxisFixed(ByVal grp As Shape, ByVal bottomRow As Boolean) As Boolean
    Dim index As Long
    Dim child As Shape
    Dim locked As MsoTriState
    Dim childCount As Long
    PPTK_GroupAxisFixed = True
    On Error Resume Next
    Err.Clear
    childCount = grp.GroupItems.Count
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    For index = 1 To childCount
        Set child = grp.GroupItems(index)
        If child Is Nothing Then
            Err.Clear
            Exit Function
        End If
        If child.Type = msoPicture Or child.Type = msoLinkedPicture Then
            Err.Clear
            Exit Function
        End If
        If child.Type = msoGroup Then
            Err.Clear
            locked = child.LockAspectRatio
            If Err.Number = 0 Then
                If locked = msoTrue Then
                    Err.Clear
                    Exit Function
                End If
            End If
            If PPTK_GroupAxisFixed(child, bottomRow) Then
                Err.Clear
                Exit Function
            End If
        Else
            Err.Clear
            locked = child.LockAspectRatio
            If Err.Number = 0 Then
                If locked = msoTrue Then
                    Err.Clear
                    Exit Function
                End If
            End If
            If PPTK_AutoSizeBlocks(child, bottomRow) Then
                Err.Clear
                Exit Function
            End If
        End If
    Next index
    PPTK_GroupAxisFixed = False
    Err.Clear
End Function

' "Resize shape to fit text" sets the height from the text. The width is set
' by the text only when the text does not wrap; wrapped text reflows to any width.
Private Function PPTK_AutoSizeBlocks(ByVal shp As Shape, ByVal bottomRow As Boolean) As Boolean
    On Error GoTo Clean
    If shp.TextFrame2.AutoSize <> msoAutoSizeShapeToFitText Then
        PPTK_AutoSizeBlocks = False
    ElseIf bottomRow Then
        PPTK_AutoSizeBlocks = True
    Else
        PPTK_AutoSizeBlocks = (shp.TextFrame2.WordWrap <> msoTrue)
    End If
    Exit Function
Clean:
    PPTK_AutoSizeBlocks = False
    Err.Clear
End Function

Private Sub PPTK_Tell(ByVal message As String)
    If Len(message) = 0 Then message = "The grid could not keep its original bounds."
    MsgBox message, vbExclamation, "PPTK"
End Sub

Private Function PPTK_SingleTable(ByRef tableShape As Shape) As Boolean
    On Error GoTo Clean
    PPTK_SingleTable = False
    ' A table selected as an object, a range of selected cells, or a text
    ' cursor inside a cell all count as the table.
    If ActiveWindow.Selection.Type <> ppSelectionShapes And ActiveWindow.Selection.Type <> ppSelectionText Then Exit Function
    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then Exit Function
    Set tableShape = ActiveWindow.Selection.ShapeRange(1)
    PPTK_SingleTable = tableShape.HasTable
Clean:
End Function

Private Function PPTK_Recognize(ByRef items() As Shape, ByRef count As Long, ByRef rowCount As Long, ByRef colCount As Long, ByRef rowOf() As Long, ByRef colOf() As Long) As Boolean
    Dim parent() As Long
    Dim index As Long
    Dim other As Long
    Dim rowTol As Single
    Dim colTol As Single
    Dim rowId() As Long
    Dim rows As Long
    Dim cols As Long
    Dim perRow() As Long
    Dim anchorX() As Single
    Dim rowY() As Single
    Dim rowRank() As Long

    PPTK_Recognize = False
    count = 0
    On Error GoTo Clean
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Function
    ReDim items(1 To ActiveWindow.Selection.ShapeRange.Count)
    For index = 1 To ActiveWindow.Selection.ShapeRange.Count
        If PPTK_GridShape(ActiveWindow.Selection.ShapeRange(index)) Then
            count = count + 1
            Set items(count) = ActiveWindow.Selection.ShapeRange(index)
        End If
    Next index
    If count < 1 Then Exit Function

    rowTol = PPTK_Tolerance(items, count, False)
    colTol = PPTK_Tolerance(items, count, True)
    ReDim parent(1 To count)
    For index = 1 To count
        parent(index) = index
    Next index
    For index = 1 To count - 1
        For other = index + 1 To count
            If Abs(PPTK_MidY(items(index)) - PPTK_MidY(items(other))) <= rowTol _
               Or PPTK_SameBand(items(index), items(other)) Then
                PPTK_Union parent, index, other
            End If
        Next other
    Next index

    ReDim rowId(1 To count)
    rows = 0
    For index = 1 To count
        rowId(index) = 0
    Next index
    For index = 1 To count
        If rowId(PPTK_Find(parent, index)) = 0 Then
            rows = rows + 1
            rowId(PPTK_Find(parent, index)) = rows
        End If
    Next index
    ' Rows are numbered top to bottom, whatever order the shapes were selected in
    ' (columns are already ordered by X), so "bottom row" is always the lowest one.
    ReDim rowY(1 To rows)
    For index = 1 To count
        rowY(rowId(PPTK_Find(parent, index))) = PPTK_MidY(items(index))
    Next index
    ReDim rowRank(1 To rows)
    For index = 1 To rows
        rowRank(index) = 1
        For other = 1 To rows
            If rowY(other) < rowY(index) Or (rowY(other) = rowY(index) And other < index) Then rowRank(index) = rowRank(index) + 1
        Next other
    Next index
    ReDim rowOf(1 To count)
    ReDim perRow(1 To rows)
    For index = 1 To count
        rowOf(index) = rowRank(rowId(PPTK_Find(parent, index)))
        perRow(rowOf(index)) = perRow(rowOf(index)) + 1
    Next index
    cols = perRow(1)
    For index = 2 To rows
        If perRow(index) <> cols Then Exit Function
    Next index
    ' Every row must share one vertical band (within the row tolerance). A tall
    ' shape that overlaps both a top and a bottom cell chains them into one "row"
    ' with no common band: that layout is ambiguous and is not a grid.
    If Not PPTK_RowsShareBand(items, count, rowOf, rows, rowTol) Then Exit Function

    ReDim colOf(1 To count)
    ReDim anchorX(1 To cols)
    PPTK_AssignColumns items, count, rowOf, rows, cols, colTol, colOf, anchorX
    If cols < 1 Then Exit Function
    rowCount = rows
    colCount = cols
    PPTK_Recognize = True
Clean:
End Function

Private Sub PPTK_AssignColumns(ByRef items() As Shape, ByVal count As Long, ByRef rowOf() As Long, ByVal rows As Long, ByRef cols As Long, ByVal colTol As Single, ByRef colOf() As Long, ByRef anchorX() As Single)
    Dim order() As Long
    Dim index As Long
    Dim position As Long
    Dim row As Long
    Dim member As Long
    Dim seen As Long
    Dim column As Long
    Dim x As Single

    ReDim order(1 To count)
    For row = 1 To rows
        seen = 0
        For index = 1 To count
            If rowOf(index) = row Then
                seen = seen + 1
                order(seen) = index
            End If
        Next index
        For index = 1 To seen - 1
            For position = index + 1 To seen
                If PPTK_MidX(items(order(position))) < PPTK_MidX(items(order(index))) Then
                    member = order(index)
                    order(index) = order(position)
                    order(position) = member
                End If
            Next position
        Next index
        If row = 1 Then
            For column = 1 To seen
                anchorX(column) = PPTK_MidX(items(order(column)))
                colOf(order(column)) = column
            Next column
        Else
            For column = 1 To seen
                x = PPTK_MidX(items(order(column)))
                If Abs(x - anchorX(column)) > colTol Then
                    cols = 0
                    Exit Sub
                End If
                colOf(order(column)) = column
            Next column
        End If
    Next row
End Sub

Private Function PPTK_GridShape(ByVal shp As Shape) As Boolean
    On Error GoTo Clean
    Select Case shp.Type
        Case msoPicture, msoLinkedPicture, msoMedia, msoOLEControlObject, _
             msoEmbeddedOLEObject, msoLinkedOLEObject, msoInk, msoInkComment, msoLine
            PPTK_GridShape = False
        Case Else
            ' Connectors join cells; they are never cells themselves.
            PPTK_GridShape = Not shp.HasTable And Not shp.HasChart And Not shp.Connector
    End Select
    Exit Function
Clean:
    PPTK_GridShape = False
End Function

Private Function PPTK_Tolerance(ByRef items() As Shape, ByVal count As Long, ByVal horizontal As Boolean) As Single
    Dim values() As Single
    Dim index As Long
    Dim median As Single
    ReDim values(1 To count)
    For index = 1 To count
        If horizontal Then
            values(index) = items(index).Width
        Else
            values(index) = items(index).Height
        End If
    Next index
    PPTK_SortNumbers values, count
    median = values((count + 1) \ 2)
    PPTK_Tolerance = 0.35 * median
    If PPTK_Tolerance < 6 Then PPTK_Tolerance = 6
    If PPTK_Tolerance > 18 Then PPTK_Tolerance = 18
End Function

' Spacing for a new Expand row or column: the median measured gap along that axis,
' zero included (touching cells stay touching; overlaps count as zero). A single
' row or column has no gap on that axis, so the other axis's spacing is used;
' GapDefault applies only when the grid has no spacing to measure at all.
Private Function PPTK_PositiveGap(ByRef items() As Shape, ByVal count As Long, ByRef rowOf() As Long, ByRef colOf() As Long, ByVal rowCount As Long, ByVal colCount As Long, ByVal vertical As Boolean) As Single
    Dim gap As Single
    If PPTK_MedianGap(items, count, rowOf, colOf, vertical, gap) Then
        PPTK_PositiveGap = gap
    ElseIf PPTK_MedianGap(items, count, rowOf, colOf, Not vertical, gap) Then
        PPTK_PositiveGap = gap
    Else
        PPTK_PositiveGap = GapDefault
    End If
End Function

Private Function PPTK_MedianGap(ByRef items() As Shape, ByVal count As Long, ByRef rowOf() As Long, ByRef colOf() As Long, ByVal vertical As Boolean, ByRef median As Single) As Boolean
    Dim gaps() As Single
    Dim found As Long
    Dim index As Long
    Dim other As Long
    Dim gap As Single
    Dim paired As Boolean
    found = 0
    ReDim gaps(1 To count)
    For index = 1 To count
        For other = 1 To count
            paired = False
            If vertical Then
                If colOf(index) = colOf(other) And rowOf(other) = rowOf(index) + 1 Then
                    gap = items(other).Top - (items(index).Top + items(index).Height)
                    paired = True
                End If
            ElseIf rowOf(index) = rowOf(other) And colOf(other) = colOf(index) + 1 Then
                gap = items(other).Left - (items(index).Left + items(index).Width)
                paired = True
            End If
            If paired Then
                If gap < 0 Then gap = 0
                found = found + 1
                gaps(found) = gap
            End If
        Next other
    Next index
    PPTK_MedianGap = (found > 0)
    If found > 0 Then
        PPTK_SortNumbers gaps, found
        median = gaps((found + 1) \ 2)
    End If
End Function

Private Sub PPTK_ShiftInside(ByRef items() As Shape, ByVal count As Long, ByRef copies() As Shape, ByVal copyCount As Long, ByVal vertical As Boolean)
    Dim index As Long
    Dim overflow As Single
    Dim limit As Single
    overflow = 0
    If vertical Then
        limit = ActivePresentation.PageSetup.SlideHeight - SlideMargin
        For index = 1 To copyCount
            If copies(index).Top + copies(index).Height - limit > overflow Then
                overflow = copies(index).Top + copies(index).Height - limit
            End If
        Next index
        If overflow <= 0 Then Exit Sub
        If items(1).Top - overflow < SlideMargin Then overflow = items(1).Top - SlideMargin
        For index = 2 To count
            If items(index).Top - overflow < SlideMargin Then overflow = items(index).Top - SlideMargin
        Next index
        If overflow <= 0 Then Exit Sub
        For index = 1 To count
            items(index).Top = items(index).Top - overflow
        Next index
        For index = 1 To copyCount
            copies(index).Top = copies(index).Top - overflow
        Next index
    Else
        limit = ActivePresentation.PageSetup.SlideWidth - SlideMargin
        For index = 1 To copyCount
            If copies(index).Left + copies(index).Width - limit > overflow Then
                overflow = copies(index).Left + copies(index).Width - limit
            End If
        Next index
        If overflow <= 0 Then Exit Sub
        If items(1).Left - overflow < SlideMargin Then overflow = items(1).Left - SlideMargin
        For index = 2 To count
            If items(index).Left - overflow < SlideMargin Then overflow = items(index).Left - SlideMargin
        Next index
        If overflow <= 0 Then Exit Sub
        For index = 1 To count
            items(index).Left = items(index).Left - overflow
        Next index
        For index = 1 To copyCount
            copies(index).Left = copies(index).Left - overflow
        Next index
    End If
End Sub

Private Sub PPTK_Union(ByRef parent() As Long, ByVal a As Long, ByVal b As Long)
    Dim rootA As Long
    Dim rootB As Long
    rootA = PPTK_Find(parent, a)
    rootB = PPTK_Find(parent, b)
    If rootA <> rootB Then parent(rootB) = rootA
End Sub

Private Function PPTK_Find(ByRef parent() As Long, ByVal index As Long) As Long
    Dim root As Long
    root = index
    Do While parent(root) <> root
        root = parent(root)
    Loop
    PPTK_Find = root
End Function

Private Sub PPTK_SortNumbers(ByRef values() As Single, ByVal count As Long)
    Dim i As Long
    Dim j As Long
    Dim swapValue As Single
    For i = 1 To count - 1
        For j = i + 1 To count
            If values(j) < values(i) Then
                swapValue = values(i)
                values(i) = values(j)
                values(j) = swapValue
            End If
        Next j
    Next i
End Sub

Private Function PPTK_DuplicateShape(ByVal shp As Shape) As Shape
    Dim result As Object
    On Error GoTo Clean
    Set result = shp.Duplicate
    If result Is Nothing Then Exit Function
    If TypeName(result) = "ShapeRange" Then
        Set PPTK_DuplicateShape = result.Item(1)
    Else
        Set PPTK_DuplicateShape = result
    End If
Clean:
    Err.Clear
End Function

' A new grid cell: a duplicate of the edge cell for its type, size, fill, line and
' font, with the text removed (also inside grouped cells). Recovery backups use
' PPTK_DuplicateShape directly and keep their text.
Private Function PPTK_NewCell(ByVal source As Shape) As Shape
    Dim cell As Shape
    Set cell = PPTK_DuplicateShape(source)
    If cell Is Nothing Then Exit Function
    PPTK_ClearCellText cell
    Set PPTK_NewCell = cell
End Function

Private Sub PPTK_ClearCellText(ByVal shp As Shape)
    Dim index As Long
    On Error Resume Next
    If shp.Type = msoGroup Then
        For index = 1 To shp.GroupItems.Count
            PPTK_ClearCellText shp.GroupItems(index)
        Next index
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame2.HasText Then shp.TextFrame2.TextRange.Text = ""
    End If
    Err.Clear
End Sub

Private Function PPTK_MidX(ByVal shp As Shape) As Single
    PPTK_MidX = shp.Left + shp.Width / 2
End Function

Private Function PPTK_MidY(ByVal shp As Shape) As Single
    PPTK_MidY = shp.Top + shp.Height / 2
End Function

' Same row when the vertical extents overlap by at least half of the shorter
' shape (top-, middle- or bottom-aligned cells of different heights). Shapes of
' separate rows do not overlap; a shape spanning two rows joins them, and the
' unequal row counts then reject the selection as not a grid.
Private Function PPTK_SameBand(ByVal a As Shape, ByVal b As Shape) As Boolean
    Dim overlap As Single
    Dim shorter As Single
    overlap = PPTK_MinSingle(a.Top + a.Height, b.Top + b.Height) - PPTK_MaxSingle(a.Top, b.Top)
    shorter = PPTK_MinSingle(a.Height, b.Height)
    PPTK_SameBand = (shorter > 0 And overlap >= 0.5 * shorter)
End Function

Private Function PPTK_RowsShareBand(ByRef items() As Shape, ByVal count As Long, ByRef rowOf() As Long, ByVal rows As Long, ByVal rowTol As Single) As Boolean
    Dim row As Long
    Dim index As Long
    Dim lowestTop As Single
    Dim highestBottom As Single
    Dim seen As Boolean
    PPTK_RowsShareBand = False
    For row = 1 To rows
        seen = False
        For index = 1 To count
            If rowOf(index) = row Then
                If Not seen Then
                    lowestTop = items(index).Top
                    highestBottom = items(index).Top + items(index).Height
                    seen = True
                Else
                    lowestTop = PPTK_MaxSingle(lowestTop, items(index).Top)
                    highestBottom = PPTK_MinSingle(highestBottom, items(index).Top + items(index).Height)
                End If
            End If
        Next index
        If lowestTop - highestBottom > rowTol Then Exit Function
    Next row
    PPTK_RowsShareBand = True
End Function

Private Function PPTK_MinSingle(ByVal a As Single, ByVal b As Single) As Single
    If a < b Then PPTK_MinSingle = a Else PPTK_MinSingle = b
End Function

Private Function PPTK_MaxSingle(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then PPTK_MaxSingle = a Else PPTK_MaxSingle = b
End Function
