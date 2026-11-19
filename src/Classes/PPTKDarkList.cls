VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPTKDarkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A dark PPTK list with the Smart Bar's thin scrollbar, used instead of MSForms ListBox / ComboBox lists
' (their scrollbars are drawn by Windows and cannot be coloured). Rows are labels in a frame; only the
' visible rows exist and are refilled when the list scrolls (the Smart Bar approach).
' Items are 1-based; ListIndex 0 means nothing selected.
' Events: Changed (the selection changed), Picked (a row was clicked), Activated (a row was double-clicked).
' Keyboard: the form passes keys to HandleKey (Up, Down, Page Up, Page Down, Home, End).
' Mouse: rows, a click on the track pages, the thumb drags. The mouse wheel is not available to MSForms
' without Win32 (the Smart Bar has the same limit).

Public Event Changed()
Public Event Picked()
Public Event Activated()

Private Const ScrollW As Single = 5
Private Const Inset As Single = 2

Private Panel As MSForms.Frame
Private WithEvents Track As MSForms.Label
Private WithEvents Thumb As MSForms.Label
Private Rows() As PPTKListRow
Private Items() As String
Private Tags() As String
Private ItemCount As Long
Private Selected As Long
Private FirstVisible As Long
Private VisibleRows As Long
Private RowH As Single
Private Hovered As Long
Private Message As String
Private Dragging As Boolean
Private DragStartY As Single
Private DragStartFirst As Long

Public Sub Init(ByVal host As Object, ByVal name As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal rowHeight As Single)
    Dim index As Long
    Dim row As MSForms.Label
    RowH = rowHeight
    Set Panel = host.Controls.Add("Forms.Frame.1", name, True)
    Panel.Caption = ""
    Panel.BackColor = PPTK_DARK_SURFACE
    Panel.BorderStyle = fmBorderStyleSingle
    Panel.BorderColor = PPTK_DARK_LINE
    Panel.SpecialEffect = fmSpecialEffectFlat
    Panel.ScrollBars = fmScrollBarsNone
    Panel.Left = x
    Panel.Top = y
    Panel.Width = w
    Panel.Height = h
    On Error Resume Next
    Panel.TabStop = False
    On Error GoTo 0
    VisibleRows = Int((Panel.InsideHeight - Inset * 2) / RowH)
    If VisibleRows < 1 Then VisibleRows = 1
    ReDim Rows(1 To VisibleRows)
    For index = 1 To VisibleRows
        Set row = Panel.Controls.Add("Forms.Label.1", name & "_row" & index, True)
        row.BackStyle = fmBackStyleOpaque
        row.BorderStyle = fmBorderStyleNone
        row.WordWrap = False
        row.Font.Name = "Segoe UI"
        row.Font.Size = 9
        row.Left = Inset
        row.Top = Inset + (index - 1) * RowH
        row.Width = Panel.InsideWidth - Inset * 2 - ScrollW - 3
        row.Height = RowH
        Set Rows(index) = New PPTKListRow
        Rows(index).Attach Me, row, index
    Next index
    Set Track = Panel.Controls.Add("Forms.Label.1", name & "_track", True)
    Set Thumb = Panel.Controls.Add("Forms.Label.1", name & "_thumb", True)
    Track.Caption = "": Track.BackStyle = fmBackStyleOpaque: Track.BackColor = PPTK_DARK_PAGE: Track.BorderStyle = fmBorderStyleNone
    Thumb.Caption = "": Thumb.BackStyle = fmBackStyleOpaque: Thumb.BackColor = PPTK_DARK_STRONG_LINE: Thumb.BorderStyle = fmBorderStyleNone
    Track.Left = Panel.InsideWidth - Inset - ScrollW
    Track.Top = Inset
    Track.Width = ScrollW
    Track.Height = Panel.InsideHeight - Inset * 2
    Thumb.Left = Track.Left
    Thumb.Width = ScrollW
    ReDim Items(0 To 0)
    ReDim Tags(0 To 0)
    ItemCount = 0
    Selected = 0
    FirstVisible = 1
    PPTK_Render
End Sub

' names / tags: 0- or 1-based arrays of strings (tags optional, same length).
Public Sub SetItems(ByVal names As Variant, Optional ByVal tagValues As Variant)
    Dim index As Long
    Dim offset As Long
    ItemCount = 0
    ReDim Items(0 To 0)
    ReDim Tags(0 To 0)
    If IsArray(names) Then
        If UBound(names) >= LBound(names) Then
            ItemCount = UBound(names) - LBound(names) + 1
            ReDim Items(1 To ItemCount)
            ReDim Tags(1 To ItemCount)
            offset = LBound(names) - 1
            For index = 1 To ItemCount
                Items(index) = CStr(names(index + offset))
                If Not IsMissing(tagValues) Then
                    If IsArray(tagValues) Then Tags(index) = CStr(tagValues(index + LBound(tagValues) - 1))
                End If
            Next index
        End If
    End If
    Selected = 0
    FirstVisible = 1
    Hovered = 0
    PPTK_Render
End Sub

Public Property Get Count() As Long
    Count = ItemCount
End Property

Public Property Get ListIndex() As Long
    ListIndex = Selected
End Property

Public Property Let ListIndex(ByVal value As Long)
    If value < 0 Or value > ItemCount Then Exit Property
    Selected = value
    PPTK_EnsureVisible
    PPTK_Render
    RaiseEvent Changed
End Property

Public Function ItemText(ByVal index As Long) As String
    If index >= 1 And index <= ItemCount Then ItemText = Items(index)
End Function

Public Function ItemTag(ByVal index As Long) As String
    If index >= 1 And index <= ItemCount Then ItemTag = Tags(index)
End Function

' Index of the item with this text (case-insensitive), 0 when none.
Public Function IndexOf(ByVal text As String) As Long
    Dim index As Long
    For index = 1 To ItemCount
        If StrComp(Items(index), text, vbTextCompare) = 0 Then IndexOf = index: Exit Function
    Next index
End Function

' A muted line shown in the first row when the list is empty ("No matching entries").
Public Property Let EmptyMessage(ByVal text As String)
    Message = text
    PPTK_Render
End Property

Public Property Get Visible() As Boolean
    Visible = Panel.Visible
End Property

Public Property Let Visible(ByVal value As Boolean)
    Panel.Visible = value
    If value Then Panel.ZOrder 0
End Property

Public Function PanelFrame() As MSForms.Frame
    Set PanelFrame = Panel
End Function

' Up, Down, Page Up, Page Down, Home, End. True when the key was used.
Public Function HandleKey(ByVal keyCode As Long) As Boolean
    Dim target As Long
    If ItemCount = 0 Then Exit Function
    Select Case keyCode
        Case vbKeyDown: target = Selected + 1
        Case vbKeyUp: target = Selected - 1
        Case vbKeyPageDown: target = Selected + VisibleRows
        Case vbKeyPageUp: target = Selected - VisibleRows
        Case vbKeyHome: target = 1
        Case vbKeyEnd: target = ItemCount
        Case Else: Exit Function
    End Select
    If target < 1 Then target = 1
    If target > ItemCount Then target = ItemCount
    HandleKey = True
    If target <> Selected Then ListIndex = target
End Function

Public Sub Leave()
    If Hovered = 0 Then Exit Sub
    Hovered = 0
    PPTK_Render
End Sub

' ---- called by PPTKListRow

Public Sub RowClick(ByVal slot As Long)
    Dim index As Long
    index = FirstVisible + slot - 1
    If index < 1 Or index > ItemCount Then Exit Sub
    If index <> Selected Then ListIndex = index
    RaiseEvent Picked
End Sub

Public Sub RowDblClick(ByVal slot As Long)
    Dim index As Long
    index = FirstVisible + slot - 1
    If index < 1 Or index > ItemCount Then Exit Sub
    If index <> Selected Then ListIndex = index
    RaiseEvent Activated
End Sub

Public Sub RowHover(ByVal slot As Long)
    Dim index As Long
    index = FirstVisible + slot - 1
    If index > ItemCount Then index = 0
    If index = Hovered Then Exit Sub
    Hovered = index
    PPTK_Render
End Sub

' ---- drawing and scrolling

Private Sub PPTK_Render()
    Dim slot As Long
    Dim index As Long
    Dim row As MSForms.Label
    For slot = 1 To VisibleRows
        Set row = Rows(slot).RowLabel
        index = FirstVisible + slot - 1
        If index <= ItemCount Then
            row.Caption = "  " & Items(index)
            row.ForeColor = PPTK_DARK_INK
            If index = Selected Then
                row.BackColor = PPTK_DARK_SELECTED
            ElseIf index = Hovered Then
                row.BackColor = PPTK_DARK_HOVER
            Else
                row.BackColor = PPTK_DARK_SURFACE
            End If
        ElseIf ItemCount = 0 And slot = 1 And Len(Message) > 0 Then
            row.Caption = "  " & Message
            row.ForeColor = PPTK_DARK_MUTED
            row.BackColor = PPTK_DARK_SURFACE
        Else
            row.Caption = ""
            row.BackColor = PPTK_DARK_SURFACE
        End If
    Next slot
    PPTK_SyncScroll
End Sub

Private Function PPTK_LastTop() As Long
    PPTK_LastTop = ItemCount - VisibleRows + 1
    If PPTK_LastTop < 1 Then PPTK_LastTop = 1
End Function

Private Sub PPTK_EnsureVisible()
    If Selected < 1 Then Exit Sub
    If Selected < FirstVisible Then FirstVisible = Selected
    If Selected > FirstVisible + VisibleRows - 1 Then FirstVisible = Selected - VisibleRows + 1
    If FirstVisible < 1 Then FirstVisible = 1
End Sub

Private Sub PPTK_SyncScroll()
    Dim travel As Single
    Dim lastTop As Long
    lastTop = PPTK_LastTop()
    If FirstVisible > lastTop Then FirstVisible = lastTop
    Track.Visible = (ItemCount > VisibleRows)
    Thumb.Visible = Track.Visible
    If Not Track.Visible Then Exit Sub
    Thumb.Height = Track.Height * VisibleRows / ItemCount
    If Thumb.Height < 16 Then Thumb.Height = 16
    travel = Track.Height - Thumb.Height
    If lastTop > 1 Then
        Thumb.Top = Track.Top + travel * (FirstVisible - 1) / (lastTop - 1)
    Else
        Thumb.Top = Track.Top
    End If
End Sub

Private Sub PPTK_ScrollTo(ByVal firstRow As Long)
    If firstRow < 1 Then firstRow = 1
    If firstRow > PPTK_LastTop() Then firstRow = PPTK_LastTop()
    If firstRow = FirstVisible Then Exit Sub
    FirstVisible = firstRow
    PPTK_Render
End Sub

Private Sub Track_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Y < Thumb.Top - Track.Top Then
        PPTK_ScrollTo FirstVisible - VisibleRows
    ElseIf Y > Thumb.Top - Track.Top + Thumb.Height Then
        PPTK_ScrollTo FirstVisible + VisibleRows
    End If
End Sub

Private Sub Thumb_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dragging = True
    DragStartY = Thumb.Top + Y
    DragStartFirst = FirstVisible
End Sub

Private Sub Thumb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim travel As Single
    If Not Dragging Or (Button And 1) = 0 Then Exit Sub
    travel = Track.Height - Thumb.Height
    If travel <= 0 Then Exit Sub
    PPTK_ScrollTo DragStartFirst + CLng((Thumb.Top + Y - DragStartY) / travel * (PPTK_LastTop() - 1))
End Sub

Private Sub Thumb_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dragging = False
End Sub
