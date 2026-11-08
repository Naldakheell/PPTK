Attribute VB_Name = "ModulePPTKGrid"
Option Explicit

' PPTK row and column commands. These are local geometric rules.
'
' A native PowerPoint table is handled only when it is the single selected
' shape. Adding uses Table.Columns.Add or Table.Rows.Add, which keeps the
' existing cells and inserts an empty row or column at the end. Deleting
' removes only that last row or column. A table with one row or one column
' is not deleted. This path is separate from the shape-grid path.
'
' Shape grids use the current shape selection only. Nothing else on the
' slide is moved or deleted. Pictures, lines, connectors, media, OLE, and
' ink are ignored. A group counts as one shape.
'
' Rows share a vertical center within Max(6, Min(18, 0.35 * median height))
' points. Columns share a horizontal center within the same rule using width.
' A grid is recognized only when every row has the same count and the columns
' line up within that tolerance. A single row or a single column is valid.
' A ragged or staggered selection does nothing, including deletion.
'
' Add Right Column duplicates each shape in the rightmost column. Add Bottom
' Row duplicates each shape in the bottom row. Duplicates keep the source
' text, size, formatting, and type. The new gap is the median positive gap
' already in the grid, or 8 points. New shapes share the outer edge of the
' source row or column. If they would leave the slide, the selected shapes
' and the copies shift back together, and stop at the slide margin. Other
' slide content stays put.
'
' Delete Right Column and Delete Bottom Row remove only that recognized edge.
' They do nothing when the grid has only one column or only one row.

Private Const GapDefault As Single = 8
Private Const SlideMargin As Single = 8

Public Sub PPTK_AddRightColumn()
    PPTK_AddAxis False
End Sub

Public Sub PPTK_AddBottomRow()
    PPTK_AddAxis True
End Sub

Public Sub PPTK_DeleteRightColumn()
    PPTK_DeleteAxis False
End Sub

Public Sub PPTK_DeleteBottomRow()
    PPTK_DeleteAxis True
End Sub

Private Sub PPTK_AddAxis(ByVal bottomRow As Boolean)
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

    If PPTK_SingleTable(tableShape) Then
        On Error GoTo TableFailed
        If bottomRow Then
            tableShape.Table.Rows.Add
        Else
            tableShape.Table.Columns.Add
        End If
        Exit Sub
TableFailed:
        Exit Sub
    End If

    If Not PPTK_Recognize(items, count, rowCount, colCount, rowOf, colOf) Then Exit Sub
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
    For index = 1 To edgeCount
        If edge(index) Is Nothing Then Exit Sub
        Set copies(index) = PPTK_DuplicateShape(edge(index))
        If copies(index) Is Nothing Then Exit Sub
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
End Sub

Private Sub PPTK_DeleteAxis(ByVal bottomRow As Boolean)
    Dim tableShape As Shape
    Dim items() As Shape
    Dim count As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim rowOf() As Long
    Dim colOf() As Long
    Dim index As Long

    If PPTK_SingleTable(tableShape) Then
        On Error GoTo TableFailed
        If bottomRow Then
            If tableShape.Table.Rows.Count <= 1 Then Exit Sub
            tableShape.Table.Rows(tableShape.Table.Rows.Count).Delete
        Else
            If tableShape.Table.Columns.Count <= 1 Then Exit Sub
            tableShape.Table.Columns(tableShape.Table.Columns.Count).Delete
        End If
        Exit Sub
TableFailed:
        Exit Sub
    End If

    If Not PPTK_Recognize(items, count, rowCount, colCount, rowOf, colOf) Then Exit Sub
    If bottomRow Then
        If rowCount <= 1 Then Exit Sub
    Else
        If colCount <= 1 Then Exit Sub
    End If
    For index = count To 1 Step -1
        If bottomRow And rowOf(index) = rowCount Then items(index).Delete
        If Not bottomRow And colOf(index) = colCount Then items(index).Delete
    Next index
End Sub

Private Function PPTK_SingleTable(ByRef tableShape As Shape) As Boolean
    On Error GoTo Clean
    PPTK_SingleTable = False
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Function
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
            If Abs(PPTK_MidY(items(index)) - PPTK_MidY(items(other))) <= rowTol Then
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
    ReDim rowOf(1 To count)
    ReDim perRow(1 To rows)
    For index = 1 To count
        rowOf(index) = rowId(PPTK_Find(parent, index))
        perRow(rowOf(index)) = perRow(rowOf(index)) + 1
    Next index
    cols = perRow(1)
    For index = 2 To rows
        If perRow(index) <> cols Then Exit Function
    Next index

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
            PPTK_GridShape = Not shp.HasTable And Not shp.HasChart
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

Private Function PPTK_PositiveGap(ByRef items() As Shape, ByVal count As Long, ByRef rowOf() As Long, ByRef colOf() As Long, ByVal rowCount As Long, ByVal colCount As Long, ByVal vertical As Boolean) As Single
    Dim gaps() As Single
    Dim found As Long
    Dim index As Long
    Dim other As Long
    Dim gap As Single
    found = 0
    ReDim gaps(1 To count)
    For index = 1 To count
        For other = 1 To count
            If vertical Then
                If colOf(index) = colOf(other) And rowOf(other) = rowOf(index) + 1 Then
                    gap = items(other).Top - (items(index).Top + items(index).Height)
                    If gap > 0.5 Then
                        found = found + 1
                        gaps(found) = gap
                    End If
                End If
            ElseIf rowOf(index) = rowOf(other) And colOf(other) = colOf(index) + 1 Then
                gap = items(other).Left - (items(index).Left + items(index).Width)
                If gap > 0.5 Then
                    found = found + 1
                    gaps(found) = gap
                End If
            End If
        Next other
    Next index
    If found = 0 Then
        PPTK_PositiveGap = GapDefault
    Else
        PPTK_SortNumbers gaps, found
        PPTK_PositiveGap = gaps((found + 1) \ 2)
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
End Function

Private Function PPTK_MidX(ByVal shp As Shape) As Single
    PPTK_MidX = shp.Left + shp.Width / 2
End Function

Private Function PPTK_MidY(ByVal shp As Shape) As Single
    PPTK_MidY = shp.Top + shp.Height / 2
End Function
