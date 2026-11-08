Attribute VB_Name = "ModulePPTKSmartAlign"
Option Explicit

' PPTK smart alignment. Auxi does not publish exact coordinates, so these
' are PPTK rules.
'
' Shared selection rules:
'   Shape selection only. Text editing, pictures, lines, connectors, tables,
'   charts, media, OLE, and ink stay where they are. A group is one item and
'   is not ungrouped. Fewer than two usable shapes does nothing.
'   Order is deterministic: position first, then shape Id.
'   Size, text, formatting, and shape type are preserved.
'
' Align Process (A P):
'   Usable autoshapes, text boxes, placeholders, freeforms, callouts, and
'   groups are arranged. Chevrons and pentagons are included with the other
'   usable shapes rather than pulled out of a mixed selection.
'   The flow is horizontal when the selection is at least as wide as it is
'   tall, otherwise vertical. Shapes are sorted along that flow.
'   The first shape keeps its leading edge. Later shapes follow it with one
'   shared gap and a shared center line. The gap is the median positive gap
'   already present, or 8 points. It is reduced toward zero if that is needed
'   to stay on the slide. Shapes are not resized. Overflow past the far edge
'   is left in place rather than squashing the shapes.
'
' Align Matrix (A X):
'   A modal grid picker opens. The default is the nearest square
'   that can hold the selection. Rows times columns must be at least the
'   shape count, and each side must be from 1 through 12.
'   Shapes are ordered by top, then left, then Id, and filled row by row.
'   Each shape keeps its size and is centered in a cell of the largest width
'   and height. The gap starts at 12 points and shrinks to 4 before the grid
'   is allowed to meet the slide margin. The grid is centered on the slide.
'
' Align Central (A Z):
'   One selected shape is left unchanged. Two or more open a radius picker,
'   in points, defaulting to the average distance from the selection centroid.
'   Shapes are ordered clockwise from the top, with Id as the tie break, then
'   placed on equal angles starting at the top. The requested radius is
'   reduced when a shape would leave the slide. Radius zero is allowed and
'   does not divide by zero.

Private Const ProcessGapDefault As Single = 8
Private Const MatrixGap As Single = 12
Private Const MatrixGapFloor As Single = 4
Private Const SlideMargin As Single = 8

Public Sub PPTK_AlignProcess()
    Dim items() As Shape
    Dim count As Long
    Dim boundsWidth As Single
    Dim boundsHeight As Single
    PPTK_CollectUsable items, count, True
    If count < 2 Then Exit Sub
    PPTK_Bounds items, count, boundsWidth, boundsHeight
    If boundsWidth >= boundsHeight Then
        PPTK_Sort items, count, True
        PPTK_Flow items, count, True
    Else
        PPTK_Sort items, count, False
        PPTK_Flow items, count, False
    End If
End Sub

Public Sub PPTK_AlignMatrix(Optional ByVal rows As Long = 0, Optional ByVal cols As Long = 0)
    Dim items() As Shape
    Dim count As Long
    Dim cellW As Single
    Dim cellH As Single
    Dim gapX As Single
    Dim gapY As Single
    Dim totalW As Single
    Dim totalH As Single
    Dim originX As Single
    Dim originY As Single
    Dim index As Long
    Dim row As Long
    Dim col As Long
    Dim slideW As Single
    Dim slideH As Single
    Dim ask As Boolean

    PPTK_CollectUsable items, count, False
    If count < 2 Then Exit Sub
    ask = (rows < 1 Or cols < 1)
    If ask Then
        cols = CLng(Int(Sqr(count)))
        If cols < 1 Then cols = 1
        If cols * cols < count Then cols = cols + 1
        rows = (count + cols - 1) \ cols
        If Not PPTKMatrixForm.PPTK_Ask(count, rows, cols) Then Exit Sub
    End If
    If rows < 1 Or cols < 1 Or rows > 12 Or cols > 12 Then Exit Sub
    If rows * cols < count Then Exit Sub

    PPTK_SortReading items, count
    cellW = 0
    cellH = 0
    For index = 1 To count
        If items(index).Width > cellW Then cellW = items(index).Width
        If items(index).Height > cellH Then cellH = items(index).Height
    Next index

    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    gapX = MatrixGap
    gapY = MatrixGap
    If cols > 1 Then
        gapX = PPTK_FitGap(slideW - (SlideMargin * 2), cols, cellW, MatrixGap, MatrixGapFloor)
    End If
    If rows > 1 Then
        gapY = PPTK_FitGap(slideH - (SlideMargin * 2), rows, cellH, MatrixGap, MatrixGapFloor)
    End If
    totalW = cols * cellW
    If cols > 1 Then totalW = totalW + (cols - 1) * gapX
    totalH = rows * cellH
    If rows > 1 Then totalH = totalH + (rows - 1) * gapY
    originX = (slideW - totalW) / 2
    originY = (slideH - totalH) / 2
    If originX < SlideMargin Then originX = SlideMargin
    If originY < SlideMargin Then originY = SlideMargin

    For index = 1 To count
        row = (index - 1) \ cols
        col = (index - 1) Mod cols
        items(index).Left = originX + col * (cellW + gapX) + (cellW - items(index).Width) / 2
        items(index).Top = originY + row * (cellH + gapY) + (cellH - items(index).Height) / 2
    Next index
End Sub

Public Sub PPTK_AlignCentral(Optional ByVal radius As Single = -1)
    Dim items() As Shape
    Dim count As Long
    Dim index As Long
    Dim cx As Single
    Dim cy As Single
    Dim fitted As Single
    Dim theta As Double
    Dim stepAngle As Double
    Dim limit As Single
    Dim ask As Boolean

    PPTK_CollectUsable items, count, False
    If count < 2 Then Exit Sub
    PPTK_Centroid items, count, cx, cy
    ask = (radius < 0)
    If ask Then
        radius = 0
        For index = 1 To count
            radius = radius + PPTK_Distance(PPTK_CenterX(items(index)), PPTK_CenterY(items(index)), cx, cy)
        Next index
        radius = radius / count
        If radius < 36 Then radius = 36
        If Not PPTKCentralForm.PPTK_Ask(radius) Then Exit Sub
    End If
    If radius < 0 Then Exit Sub

    PPTK_SortClockwise items, count, cx, cy
    stepAngle = 6.28318530717959 / count
    fitted = radius
    For index = 1 To count
        theta = -1.5707963267949 + (index - 1) * stepAngle
        limit = PPTK_RadiusLimit(cx, cy, items(index).Width, items(index).Height, theta)
        If limit < fitted Then fitted = limit
    Next index
    If fitted < 0 Then fitted = 0

    For index = 1 To count
        theta = -1.5707963267949 + (index - 1) * stepAngle
        items(index).Left = cx + fitted * Cos(theta) - items(index).Width / 2
        items(index).Top = cy + fitted * Sin(theta) - items(index).Height / 2
        PPTK_ClampToSlide items(index)
    Next index
End Sub

Private Sub PPTK_CollectUsable(ByRef items() As Shape, ByRef count As Long, ByVal processOnly As Boolean)
    Dim picked As Shape
    Dim index As Long
    count = 0
    On Error GoTo Clean
    If ActiveWindow Is Nothing Then Exit Sub
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    ReDim items(1 To ActiveWindow.Selection.ShapeRange.Count)
    For index = 1 To ActiveWindow.Selection.ShapeRange.Count
        Set picked = ActiveWindow.Selection.ShapeRange(index)
        If PPTK_Usable(picked, processOnly) Then
            count = count + 1
            Set items(count) = picked
        End If
    Next index
    Exit Sub
Clean:
    count = 0
End Sub

Private Function PPTK_Usable(ByVal shp As Shape, ByVal processOnly As Boolean) As Boolean
    On Error GoTo Clean
    If processOnly Then
        Select Case shp.Type
            Case msoAutoShape, msoTextBox, msoPlaceholder, msoFreeform, msoCallout, msoGroup
                PPTK_Usable = True
            Case Else
                PPTK_Usable = False
        End Select
    Else
        Select Case shp.Type
            Case msoPicture, msoLinkedPicture, msoMedia, msoOLEControlObject, _
                 msoEmbeddedOLEObject, msoLinkedOLEObject, msoInk, msoInkComment, _
                 msoLine, msoTextEffect
                PPTK_Usable = False
            Case Else
                If shp.HasTable Or shp.HasChart Then
                    PPTK_Usable = False
                Else
                    PPTK_Usable = True
                End If
        End Select
    End If
    If shp.Type = msoLine Then PPTK_Usable = False
    Exit Function
Clean:
    PPTK_Usable = False
End Function

Private Sub PPTK_Bounds(ByRef items() As Shape, ByVal count As Long, ByRef width As Single, ByRef height As Single)
    Dim index As Long
    Dim minX As Single
    Dim minY As Single
    Dim maxX As Single
    Dim maxY As Single
    minX = items(1).Left
    minY = items(1).Top
    maxX = items(1).Left + items(1).Width
    maxY = items(1).Top + items(1).Height
    For index = 2 To count
        If items(index).Left < minX Then minX = items(index).Left
        If items(index).Top < minY Then minY = items(index).Top
        If items(index).Left + items(index).Width > maxX Then maxX = items(index).Left + items(index).Width
        If items(index).Top + items(index).Height > maxY Then maxY = items(index).Top + items(index).Height
    Next index
    width = maxX - minX
    height = maxY - minY
End Sub

Private Sub PPTK_Flow(ByRef items() As Shape, ByVal count As Long, ByVal horizontal As Boolean)
    Dim index As Long
    Dim gap As Single
    Dim cursor As Single
    Dim anchorCenter As Single
    Dim span As Single
    Dim available As Single
    Dim widths As Single

    gap = PPTK_MedianGap(items, count, horizontal)
    If gap <= 0 Then gap = ProcessGapDefault
    widths = 0
    For index = 1 To count
        If horizontal Then
            widths = widths + items(index).Width
        Else
            widths = widths + items(index).Height
        End If
    Next index
    If horizontal Then
        cursor = items(1).Left
        available = ActivePresentation.PageSetup.SlideWidth - SlideMargin - cursor
    Else
        cursor = items(1).Top
        available = ActivePresentation.PageSetup.SlideHeight - SlideMargin - cursor
    End If
    span = widths + gap * (count - 1)
    If span > available And count > 1 Then
        gap = (available - widths) / (count - 1)
        If gap < 0 Then gap = 0
    End If

    If horizontal Then
        anchorCenter = PPTK_CenterY(items(1))
        For index = 1 To count
            items(index).Left = cursor
            items(index).Top = anchorCenter - items(index).Height / 2
            cursor = cursor + items(index).Width + gap
        Next index
    Else
        anchorCenter = PPTK_CenterX(items(1))
        For index = 1 To count
            items(index).Top = cursor
            items(index).Left = anchorCenter - items(index).Width / 2
            cursor = cursor + items(index).Height + gap
        Next index
    End If
End Sub

Private Function PPTK_MedianGap(ByRef items() As Shape, ByVal count As Long, ByVal horizontal As Boolean) As Single
    Dim gaps() As Single
    Dim found As Long
    Dim index As Long
    Dim gap As Single
    found = 0
    If count < 2 Then
        PPTK_MedianGap = 0
        Exit Function
    End If
    ReDim gaps(1 To count - 1)
    For index = 1 To count - 1
        If horizontal Then
            gap = items(index + 1).Left - (items(index).Left + items(index).Width)
        Else
            gap = items(index + 1).Top - (items(index).Top + items(index).Height)
        End If
        If gap > 0.5 Then
            found = found + 1
            gaps(found) = gap
        End If
    Next index
    If found = 0 Then
        PPTK_MedianGap = 0
    Else
        PPTK_SortNumbers gaps, found
        PPTK_MedianGap = gaps((found + 1) \ 2)
    End If
End Function

Private Function PPTK_FitGap(ByVal available As Single, ByVal cells As Long, ByVal cell As Single, ByVal preferred As Single, ByVal floor As Single) As Single
    Dim needed As Single
    needed = preferred
    If cells > 1 Then
        If cells * cell + (cells - 1) * needed > available Then
            needed = (available - cells * cell) / (cells - 1)
        End If
    End If
    If needed < floor Then needed = floor
    If needed < 0 Then needed = 0
    PPTK_FitGap = needed
End Function

Private Sub PPTK_Centroid(ByRef items() As Shape, ByVal count As Long, ByRef cx As Single, ByRef cy As Single)
    Dim index As Long
    cx = 0
    cy = 0
    For index = 1 To count
        cx = cx + PPTK_CenterX(items(index))
        cy = cy + PPTK_CenterY(items(index))
    Next index
    cx = cx / count
    cy = cy / count
End Sub

Private Function PPTK_RadiusLimit(ByVal cx As Single, ByVal cy As Single, ByVal width As Single, ByVal height As Single, ByVal theta As Double) As Single
    Dim cosine As Double
    Dim sine As Double
    Dim limit As Single
    Dim room As Single
    cosine = Cos(theta)
    sine = Sin(theta)
    limit = 100000
    If cosine > 0.001 Then
        room = (ActivePresentation.PageSetup.SlideWidth - SlideMargin - width / 2 - cx) / cosine
        If room < limit Then limit = room
    ElseIf cosine < -0.001 Then
        room = (SlideMargin + width / 2 - cx) / cosine
        If room < limit Then limit = room
    End If
    If sine > 0.001 Then
        room = (ActivePresentation.PageSetup.SlideHeight - SlideMargin - height / 2 - cy) / sine
        If room < limit Then limit = room
    ElseIf sine < -0.001 Then
        room = (SlideMargin + height / 2 - cy) / sine
        If room < limit Then limit = room
    End If
    If limit < 0 Then limit = 0
    If limit > 100000 Then limit = 0
    PPTK_RadiusLimit = limit
End Function

Private Sub PPTK_ClampToSlide(ByVal shp As Shape)
    Dim slideW As Single
    Dim slideH As Single
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    If shp.Left < SlideMargin Then shp.Left = SlideMargin
    If shp.Top < SlideMargin Then shp.Top = SlideMargin
    If shp.Left + shp.Width > slideW - SlideMargin And shp.Width < slideW Then
        shp.Left = slideW - SlideMargin - shp.Width
    End If
    If shp.Top + shp.Height > slideH - SlideMargin And shp.Height < slideH Then
        shp.Top = slideH - SlideMargin - shp.Height
    End If
End Sub

Private Sub PPTK_Sort(ByRef items() As Shape, ByVal count As Long, ByVal horizontal As Boolean)
    Dim i As Long
    Dim j As Long
    Dim keyA As Single
    Dim keyB As Single
    Dim swapItem As Shape
    For i = 1 To count - 1
        For j = i + 1 To count
            If horizontal Then
                keyA = PPTK_CenterX(items(i))
                keyB = PPTK_CenterX(items(j))
            Else
                keyA = PPTK_CenterY(items(i))
                keyB = PPTK_CenterY(items(j))
            End If
            If keyB < keyA - 0.5 Or (Abs(keyA - keyB) <= 0.5 And items(j).Id < items(i).Id) Then
                Set swapItem = items(i)
                Set items(i) = items(j)
                Set items(j) = swapItem
            End If
        Next j
    Next i
End Sub

Private Sub PPTK_SortReading(ByRef items() As Shape, ByVal count As Long)
    Dim i As Long
    Dim j As Long
    Dim swapItem As Shape
    Dim move As Boolean
    For i = 1 To count - 1
        For j = i + 1 To count
            move = False
            If items(j).Top < items(i).Top - 0.5 Then
                move = True
            ElseIf Abs(items(j).Top - items(i).Top) <= 0.5 Then
                If items(j).Left < items(i).Left - 0.5 Then
                    move = True
                ElseIf Abs(items(j).Left - items(i).Left) <= 0.5 And items(j).Id < items(i).Id Then
                    move = True
                End If
            End If
            If move Then
                Set swapItem = items(i)
                Set items(i) = items(j)
                Set items(j) = swapItem
            End If
        Next j
    Next i
End Sub

Private Sub PPTK_SortClockwise(ByRef items() As Shape, ByVal count As Long, ByVal cx As Single, ByVal cy As Single)
    Dim i As Long
    Dim j As Long
    Dim swapItem As Shape
    Dim angleA As Double
    Dim angleB As Double
    For i = 1 To count - 1
        For j = i + 1 To count
            angleA = PPTK_Clockwise(items(i), cx, cy)
            angleB = PPTK_Clockwise(items(j), cx, cy)
            If angleB < angleA - 0.001 Or (Abs(angleA - angleB) <= 0.001 And items(j).Id < items(i).Id) Then
                Set swapItem = items(i)
                Set items(i) = items(j)
                Set items(j) = swapItem
            End If
        Next j
    Next i
End Sub

Private Function PPTK_Clockwise(ByVal shp As Shape, ByVal cx As Single, ByVal cy As Single) As Double
    Dim dx As Double
    Dim dy As Double
    Dim deg As Double
    dx = PPTK_CenterX(shp) - cx
    dy = PPTK_CenterY(shp) - cy
    If Abs(dx) < 0.01 And Abs(dy) < 0.01 Then
        PPTK_Clockwise = 0
        Exit Function
    End If
    deg = Atn2(dy, dx) * 180# / 3.14159265358979
    PPTK_Clockwise = PPTK_Modulo(90# - deg, 360#)
End Function

Private Function PPTK_Modulo(ByVal value As Double, ByVal span As Double) As Double
    PPTK_Modulo = value - span * Int(value / span)
    If PPTK_Modulo < 0 Then PPTK_Modulo = PPTK_Modulo + span
End Function

Private Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atn2 = Atn(y / x)
    ElseIf x < 0 And y >= 0 Then
        Atn2 = Atn(y / x) + 3.14159265358979
    ElseIf x < 0 And y < 0 Then
        Atn2 = Atn(y / x) - 3.14159265358979
    ElseIf y > 0 Then
        Atn2 = 1.5707963267949
    ElseIf y < 0 Then
        Atn2 = -1.5707963267949
    Else
        Atn2 = 0
    End If
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

Private Function PPTK_CenterX(ByVal shp As Shape) As Single
    PPTK_CenterX = shp.Left + shp.Width / 2
End Function

Private Function PPTK_CenterY(ByVal shp As Shape) As Single
    PPTK_CenterY = shp.Top + shp.Height / 2
End Function

Private Function PPTK_Distance(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Single
    PPTK_Distance = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function
