VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTKSmartBarForm 
   Caption         =   "Smart Bar"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "PPTKSmartBarForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PPTKSmartBarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Compact command palette. Layout from top to bottom:
'   search icon, search field, ESC  |  divider  |  section name and count  |
'   four command rows with a scrollbar  |  divider  |  key hints.
' The window has one fixed size. With an empty search the list is one
' continuous list: 3-5 commands for the current PowerPoint selection (no
' selection, one shape, several shapes, a mixed selection, a group, a table,
' or text editing), a thin divider, then every other command in registry
' order, each command once. Typing searches the whole registry; clearing the
' search restores the list from the top. Rows are a fixed pool of label controls that
' are refilled while scrolling, so typing never creates or deletes controls.
' Row icons are BMP pictures held by hidden Image controls in this form
' (see PPTK_CommandIcon). The search icon is a Segoe Fluent Icons glyph.
' No Win32 API is used.
' The mouse wheel is not available to MSForms without Win32 calls, so the
' list scrolls with a slim scrollbar (two labels: click the track to page,
' drag the thumb), Up/Down, Page Up/Page Down, and Ctrl+Home/End.
' A row runs the token stored in ModulePPTKSmartBar through PPTK_DispatchShortcut.
' Sizes are VBA points.

Private Const ColorPage As Long = 2103313       ' #111820
Private Const ColorField As Long = 3549980      ' #1C2B36
Private Const ColorSelected As Long = 5522213   ' #254354
Private Const ColorAccent As Long = 15514147    ' #23BAEC
Private Const ColorInk As Long = 16381424       ' #F0F5F9
Private Const ColorMuted As Long = 13484463     ' #AFC1CD
Private Const ColorLine As Long = 5457460       ' #344653
Private Const ColorThumb As Long = 6706494      ' #3E5566
Private Const ClientW As Single = 340
Private Const Pad As Single = 14
Private Const RowH As Single = 30
Private Const VisibleRows As Long = 4
Private Const ScrollW As Single = 4
Private Const KeyW As Single = 76
Private Const IconW As Single = 18

Private WithEvents ScrollTrack As MSForms.Label
Private WithEvents ScrollThumb As MSForms.Label
Private Dragging As Boolean
Private DragStartY As Single
Private DragStartFirst As Long
Private RowHooks As Collection
Private RowBg() As MSForms.Label
Private RowName() As MSForms.Label
Private RowKey() As MSForms.Label
Private RowIcon() As MSForms.Image
Private ResultNames() As String
Private ResultKeys() As String
Private ResultTokens() As String
Private ResultIcons() As String
Private SelectedShapes() As Object
Private PoolCount As Long
Private VisibleCount As Long
Private ResultCount As Long
Private TotalCount As Long
Private SelectedIndex As Long
Private FirstVisible As Long
Private SelectedCount As Long
Private Busy As Boolean
Private Painting As Boolean
Private ClientH As Single
Private IconFont As String
Private ContextKind As String
Private SuggestCount As Long
Private RowDivider As MSForms.Label

Private Sub UserForm_Initialize()
    Set RowHooks = New Collection
    ' Segoe Fluent Icons ships with Windows 11. Windows 10 has the same codes in Segoe MDL2 Assets.
    IconFont = "Segoe MDL2 Assets"
    On Error Resume Next
    If Len(Dir$(Environ$("WINDIR") & "\Fonts\SegoeIcons.ttf")) > 0 Then IconFont = "Segoe Fluent Icons"
    On Error GoTo 0
    TotalCount = PPTK_CommandCount
    PPTK_ApplyChrome
    PPTK_BuildPool
    PPTK_CaptureSelection
    ContextKind = PPTK_DetectContext()
    txtFilter.Text = ""
    PPTK_Refill
End Sub

Private Sub UserForm_Activate()
    Dim kind As String
    PPTK_Layout
    On Error Resume Next
    txtFilter.SetFocus
    On Error GoTo 0
    PPTK_CaptureSelection
    ' A selection changed in PowerPoint while the palette was open updates the
    ' suggestions, but only while they are shown.
    kind = PPTK_DetectContext()
    If kind <> ContextKind Then
        ContextKind = kind
        If Len(PPTK_Normalize(txtFilter.Text)) = 0 Then PPTK_Refill
    End If
End Sub

' Called by PPTK_ShowSmartBar before the palette is shown. Reads the
' PowerPoint selection once (read-only) and shows the list from the top.
Public Sub PPTK_Open()
    PPTK_CaptureSelection
    ContextKind = PPTK_DetectContext()
    If Len(txtFilter.Text) > 0 Then
        txtFilter.Text = ""
    Else
        PPTK_Refill
    End If
End Sub

' The X button hides the palette. Unload from code (PPTK_ShowSmartBar opens
' a fresh palette each time) is allowed.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub txtFilter_Change()
    lblCue.Visible = (Len(txtFilter.Text) = 0)
    PPTK_Refill
End Sub

Private Sub txtFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    PPTK_HandleKey KeyCode, Shift
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    PPTK_HandleKey KeyCode, Shift
End Sub

Private Sub fraResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    PPTK_HandleKey KeyCode, Shift
End Sub

Private Sub ScrollTrack_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Y < ScrollThumb.Top - ScrollTrack.Top Then
        PPTK_ScrollTo FirstVisible - VisibleRows
    ElseIf Y > ScrollThumb.Top - ScrollTrack.Top + ScrollThumb.Height Then
        PPTK_ScrollTo FirstVisible + VisibleRows
    End If
    PPTK_FocusSearch
End Sub

Private Sub ScrollThumb_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dragging = True
    DragStartY = ScrollThumb.Top + Y
    DragStartFirst = FirstVisible
End Sub

Private Sub ScrollThumb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim travel As Single
    If Not Dragging Or (Button And 1) = 0 Then Exit Sub
    travel = ScrollTrack.Height - ScrollThumb.Height
    If travel <= 0 Then Exit Sub
    PPTK_ScrollTo DragStartFirst + CLng((ScrollThumb.Top + Y - DragStartY) / travel * (PPTK_LastTop() - 1))
End Sub

Private Sub ScrollThumb_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dragging = False
    PPTK_FocusSearch
End Sub

Private Sub lblCue_Click()
    PPTK_FocusSearch
End Sub

Private Sub lblField_Click()
    PPTK_FocusSearch
End Sub

Private Sub lblSearchEnter_Click()
    PPTK_FocusSearch
End Sub

Private Sub lblSearchIcon_Click()
    PPTK_FocusSearch
End Sub

Private Sub lblEsc_Click()
    Me.Hide
End Sub

Private Sub fraResults_Click()
    PPTK_FocusSearch
End Sub

Private Sub PPTK_FocusSearch()
    On Error Resume Next
    txtFilter.SetFocus
End Sub

Public Sub PPTK_RowHover(ByVal rowIndex As Long)
    Dim resultIndex As Long
    If Painting Then Exit Sub
    If rowIndex < 1 Or rowIndex > VisibleCount Then Exit Sub
    resultIndex = FirstVisible + rowIndex - 1
    If resultIndex = SelectedIndex Then Exit Sub
    PPTK_SelectRow resultIndex, False
End Sub

Public Sub PPTK_RowClick(ByVal rowIndex As Long)
    Dim resultIndex As Long
    If rowIndex < 1 Or rowIndex > VisibleCount Then Exit Sub
    resultIndex = FirstVisible + rowIndex - 1
    PPTK_SelectRow resultIndex, False
    PPTK_RunSelected
End Sub

Private Sub PPTK_HandleKey(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim ctrlDown As Boolean
    ctrlDown = (Shift And 2) <> 0
    Select Case KeyCode
        Case vbKeyEscape
            KeyCode = 0
            Me.Hide
        Case vbKeyReturn
            KeyCode = 0
            PPTK_RunSelected
        Case vbKeyDown
            KeyCode = 0
            If SelectedIndex < ResultCount Then PPTK_SelectRow SelectedIndex + 1, True
        Case vbKeyUp
            KeyCode = 0
            If SelectedIndex > 1 Then PPTK_SelectRow SelectedIndex - 1, True
        Case vbKeyPageDown
            KeyCode = 0
            If ResultCount > 0 Then PPTK_SelectRow PPTK_Clamp(SelectedIndex + VisibleRows, 1, ResultCount), True
        Case vbKeyPageUp
            KeyCode = 0
            If ResultCount > 0 Then PPTK_SelectRow PPTK_Clamp(SelectedIndex - VisibleRows, 1, ResultCount), True
        Case vbKeyHome
            If ctrlDown And ResultCount > 0 Then
                KeyCode = 0
                PPTK_SelectRow 1, True
            End If
        Case vbKeyEnd
            If ctrlDown And ResultCount > 0 Then
                KeyCode = 0
                PPTK_SelectRow ResultCount, True
            End If
    End Select
End Sub

Private Function PPTK_Clamp(ByVal value As Long, ByVal low As Long, ByVal high As Long) As Long
    If value < low Then value = low
    If value > high Then value = high
    PPTK_Clamp = value
End Function

Private Sub PPTK_ApplyChrome()
    Me.Caption = "PPTK Smart Bar"
    Me.BackColor = ColorPage
    Me.ForeColor = ColorPage
    Me.Font.Name = "Segoe UI"

    ' Controls kept from the earlier design that this layout does not use.
    lblBadge.Visible = False
    lblBrand.Visible = False
    lblSearchEnter.Visible = False
    lblHint.Visible = False
    lblField.Visible = False

    PPTK_Font lblSearchIcon, 12, False, ColorAccent
    lblSearchIcon.Font.Name = IconFont
    lblSearchIcon.Caption = ChrW$(&HE721)
    lblSearchIcon.TextAlign = fmTextAlignCenter
    lblSearchIcon.BackStyle = fmBackStyleTransparent

    txtFilter.BackColor = ColorPage
    txtFilter.ForeColor = ColorInk
    txtFilter.BorderStyle = fmBorderStyleNone
    txtFilter.SpecialEffect = fmSpecialEffectFlat
    txtFilter.Font.Name = "Segoe UI"
    txtFilter.Font.Size = 12
    txtFilter.Font.Weight = 400
    txtFilter.Multiline = False
    txtFilter.EnterKeyBehavior = False
    txtFilter.TextAlign = fmTextAlignLeft

    PPTK_Font lblCue, 12, False, ColorMuted
    lblCue.Caption = "Search " & TotalCount & " commands..."
    lblCue.TextAlign = fmTextAlignLeft
    lblCue.BackStyle = fmBackStyleTransparent
    lblCue.BorderStyle = fmBorderStyleNone

    PPTK_Font lblEsc, 8, False, ColorMuted
    lblEsc.Caption = "ESC"
    lblEsc.TextAlign = fmTextAlignRight
    lblEsc.BackStyle = fmBackStyleTransparent

    PPTK_Line lblAccent
    PPTK_Line lblFooterLine

    ' Not bold: MSForms shares font objects between these labels, so a bold
    ' label would turn every row bold. Capitals set the section apart.
    PPTK_Font lblSection, 7.5, False, ColorMuted
    lblSection.TextAlign = fmTextAlignLeft
    lblSection.BackStyle = fmBackStyleTransparent

    PPTK_Font lblTitle, 8, False, ColorMuted
    lblTitle.TextAlign = fmTextAlignRight
    lblTitle.BackStyle = fmBackStyleTransparent

    fraResults.BackColor = ColorPage
    fraResults.BorderStyle = fmBorderStyleNone
    fraResults.SpecialEffect = fmSpecialEffectFlat
    fraResults.Caption = ""
    fraResults.ScrollBars = fmScrollBarsNone
    fraResults.KeepScrollBarsVisible = fmScrollBarsNone
    On Error Resume Next
    fraResults.TabStop = False
    On Error GoTo 0

    PPTK_Font lblEmpty, 10, False, ColorMuted
    lblEmpty.Caption = "No commands found"
    lblEmpty.TextAlign = fmTextAlignCenter
    lblEmpty.BackStyle = fmBackStyleTransparent
    lblEmpty.BorderStyle = fmBorderStyleNone
    lblEmpty.Visible = False

    PPTK_Font lblNavigate, 8, False, ColorMuted
    lblNavigate.Caption = ChrW$(&H2191) & ChrW$(&H2193) & "  Navigate"
    lblNavigate.TextAlign = fmTextAlignLeft
    lblNavigate.BackStyle = fmBackStyleTransparent

    PPTK_Font lblRun, 8, False, ColorMuted
    lblRun.Caption = "Enter  to run"
    lblRun.TextAlign = fmTextAlignRight
    lblRun.BackStyle = fmBackStyleTransparent

    Set ScrollTrack = Me.Controls.Add("Forms.Label.1", "dyn_track", True)
    Set ScrollThumb = Me.Controls.Add("Forms.Label.1", "dyn_thumb", True)
    ScrollTrack.Caption = ""
    ScrollTrack.BackStyle = fmBackStyleOpaque
    ScrollTrack.BackColor = ColorField
    ScrollTrack.BorderStyle = fmBorderStyleNone
    ScrollThumb.Caption = ""
    ScrollThumb.BackStyle = fmBackStyleOpaque
    ScrollThumb.BackColor = ColorThumb
    ScrollThumb.BorderStyle = fmBorderStyleNone
End Sub

Private Sub PPTK_Line(ByVal target As MSForms.Label)
    target.Caption = ""
    target.BackColor = ColorLine
    target.BackStyle = fmBackStyleOpaque
    target.BorderStyle = fmBorderStyleNone
    target.Visible = True
End Sub

Private Sub PPTK_Layout()
    Dim extraW As Single
    Dim extraH As Single
    Dim inner As Single
    Dim y As Single

    inner = ClientW - (Pad * 2)

    ' Search row
    y = 10
    lblSearchIcon.Left = Pad - 2
    lblSearchIcon.Top = y + 3
    lblSearchIcon.Width = IconW
    lblSearchIcon.Height = 18
    lblEsc.Width = 30
    lblEsc.Height = 14
    lblEsc.Left = ClientW - Pad - lblEsc.Width
    lblEsc.Top = y + 7
    txtFilter.Left = Pad + IconW + 4
    txtFilter.Top = y
    txtFilter.Width = lblEsc.Left - txtFilter.Left - 8
    txtFilter.Height = 24
    lblCue.Left = txtFilter.Left + 8
    lblCue.Top = y + 1
    lblCue.Width = txtFilter.Width - 8
    lblCue.Height = 18

    lblAccent.Left = Pad
    lblAccent.Top = 42
    lblAccent.Width = inner
    lblAccent.Height = 0.75

    ' Section row
    lblSection.Left = Pad
    lblSection.Top = 52
    lblSection.Width = inner / 2
    lblSection.Height = 12
    lblTitle.Width = inner / 2
    lblTitle.Left = ClientW - Pad - lblTitle.Width
    lblTitle.Top = 52
    lblTitle.Height = 12

    ' Results and scrollbar
    fraResults.Left = Pad - 4
    fraResults.Top = 70
    fraResults.Width = inner + 4 - ScrollW - 6
    fraResults.Height = RowH * VisibleRows
    ScrollTrack.Left = ClientW - Pad - ScrollW
    ScrollTrack.Top = fraResults.Top + 2
    ScrollTrack.Width = ScrollW
    ScrollTrack.Height = fraResults.Height - 4
    ScrollThumb.Left = ScrollTrack.Left
    ScrollThumb.Width = ScrollW

    lblEmpty.Left = fraResults.Left
    lblEmpty.Top = fraResults.Top + (fraResults.Height - 20) / 2
    lblEmpty.Width = fraResults.Width
    lblEmpty.Height = 20

    ' Footer
    lblFooterLine.Left = Pad
    lblFooterLine.Top = fraResults.Top + fraResults.Height + 8
    lblFooterLine.Width = inner
    lblFooterLine.Height = 0.75
    lblNavigate.Left = Pad
    lblNavigate.Top = lblFooterLine.Top + 8
    lblNavigate.Width = inner / 2
    lblNavigate.Height = 14
    lblRun.Width = inner / 2
    lblRun.Left = ClientW - Pad - lblRun.Width
    lblRun.Top = lblNavigate.Top
    lblRun.Height = 14
    ClientH = lblNavigate.Top + lblNavigate.Height + 10

    ' Window size: client area plus the frame Office adds at the current scaling.
    extraW = Me.Width - Me.InsideWidth
    extraH = Me.Height - Me.InsideHeight
    If extraW < 0 Or extraW > 48 Then extraW = 0
    If extraH < 0 Or extraH > 80 Then extraH = 0
    ' Office snaps the window to whole pixels, so re-applying the size on every
    ' Activate drifts by a fraction of a point. Only resize when really off.
    If Abs(Me.Width - (ClientW + extraW)) >= 1 Then Me.Width = ClientW + extraW
    If Abs(Me.Height - (ClientH + extraH)) >= 1 Then Me.Height = ClientH + extraH

    txtFilter.ZOrder 0
    lblCue.ZOrder 0
    lblEmpty.ZOrder 0
    PPTK_Reflow
    PPTK_SyncScroll
End Sub

Private Sub PPTK_Font(ByVal target As MSForms.Label, ByVal size As Single, ByVal isBold As Boolean, ByVal color As Long)
    target.Font.Name = "Segoe UI"
    target.Font.Size = size
    ' Weight only. On these labels MSForms turns Bold = False into weight 700.
    If isBold Then target.Font.Weight = 600 Else target.Font.Weight = 400
    target.ForeColor = color
End Sub

Private Sub PPTK_BuildPool()
    Dim index As Long
    If PoolCount >= VisibleRows Then Exit Sub
    Set RowHooks = New Collection
    ReDim RowBg(1 To VisibleRows)
    ReDim RowName(1 To VisibleRows)
    ReDim RowKey(1 To VisibleRows)
    ReDim RowIcon(1 To VisibleRows)
    For index = 1 To VisibleRows
        Set RowBg(index) = PPTK_MakeLabel("dyn_bg_" & Format$(index, "00"))
        Set RowIcon(index) = PPTK_MakeImage("dyn_img_" & Format$(index, "00"))
        Set RowName(index) = PPTK_MakeLabel("dyn_name_" & Format$(index, "00"))
        Set RowKey(index) = PPTK_MakeLabel("dyn_key_" & Format$(index, "00"))
        PPTK_PrepareLabel RowBg(index)
        RowIcon(index).BorderStyle = fmBorderStyleNone
        RowIcon(index).BackStyle = fmBackStyleTransparent
        RowIcon(index).PictureSizeMode = fmPictureSizeModeZoom
        RowIcon(index).PictureAlignment = fmPictureAlignmentCenter
        RowIcon(index).Visible = False
        PPTK_PrepareLabel RowName(index)
        PPTK_Font RowName(index), 10.5, False, ColorInk
        PPTK_PrepareLabel RowKey(index)
        PPTK_Font RowKey(index), 8.5, False, ColorMuted
        RowKey(index).TextAlign = fmTextAlignRight
        PPTK_Bind RowBg(index), index
        PPTK_BindImage RowIcon(index), index
        PPTK_Bind RowName(index), index
        PPTK_Bind RowKey(index), index
    Next index
    ' Hairline between the suggestions and the other commands. Drawn over the
    ' top edge of a row, so it takes no row of its own.
    Set RowDivider = PPTK_MakeLabel("dyn_divider")
    PPTK_PrepareLabel RowDivider
    RowDivider.BackColor = ColorLine
    PoolCount = VisibleRows
End Sub

Private Function PPTK_MakeLabel(ByVal controlName As String) As MSForms.Label
    Dim existing As MSForms.Control
    On Error Resume Next
    Set existing = fraResults.Controls(controlName)
    On Error GoTo 0
    If existing Is Nothing Then
        Set PPTK_MakeLabel = fraResults.Controls.Add("Forms.Label.1", controlName, True)
    Else
        Set PPTK_MakeLabel = existing
    End If
End Function

Private Function PPTK_MakeImage(ByVal controlName As String) As MSForms.Image
    Dim existing As MSForms.Control
    On Error Resume Next
    Set existing = fraResults.Controls(controlName)
    On Error GoTo 0
    If existing Is Nothing Then
        Set PPTK_MakeImage = fraResults.Controls.Add("Forms.Image.1", controlName, True)
    Else
        Set PPTK_MakeImage = existing
    End If
End Function

Private Sub PPTK_PrepareLabel(ByVal labelControl As MSForms.Label)
    labelControl.Caption = ""
    labelControl.BorderStyle = fmBorderStyleNone
    labelControl.SpecialEffect = fmSpecialEffectFlat
    labelControl.BackStyle = fmBackStyleOpaque
    labelControl.BackColor = ColorPage
    labelControl.WordWrap = False
    labelControl.AutoSize = False
    labelControl.TextAlign = fmTextAlignLeft
    labelControl.Visible = False
End Sub

Private Sub PPTK_Bind(ByVal labelControl As MSForms.Label, ByVal index As Long)
    Dim hook As PPTKSmartBarRow
    Set hook = New PPTKSmartBarRow
    hook.Initialize labelControl, Me, index
    RowHooks.Add hook
End Sub

Private Sub PPTK_BindImage(ByVal imageControl As MSForms.Image, ByVal index As Long)
    Dim hook As PPTKSmartBarRow
    Set hook = New PPTKSmartBarRow
    hook.InitializeImage imageControl, Me, index
    RowHooks.Add hook
End Sub

' Search: case-insensitive, extra spaces ignored. Ranks:
' 0 exact name or shortcut, 1 name starts with the query, 2 name contains it,
' 3 an alias word starts with it, 4 the shortcut contains it. Ties keep the
' registry order. Search always covers the whole registry.
' Empty search: the suggestions for the current context, then every other
' command in registry order. Each command appears once.
Private Sub PPTK_Refill()
    Dim index As Long
    Dim commandName As String
    Dim shortcut As String
    Dim token As String
    Dim aliases As String
    Dim query As String
    Dim ranks() As Long
    Dim rank As Long
    Dim picks() As String
    Dim used() As Boolean

    If PoolCount < 1 Then PPTK_BuildPool
    Painting = True
    query = PPTK_Normalize(txtFilter.Text)
    ReDim ResultNames(1 To TotalCount)
    ReDim ResultKeys(1 To TotalCount)
    ReDim ResultTokens(1 To TotalCount)
    ReDim ResultIcons(1 To TotalCount)
    ResultCount = 0
    SuggestCount = 0

    If Len(query) > 0 Then
        ReDim ranks(1 To TotalCount)
        For index = 1 To TotalCount
            PPTK_CommandAt index, commandName, shortcut, token, aliases
            ranks(index) = PPTK_Rank(query, commandName, shortcut, aliases)
        Next index
        For rank = 0 To 4
            For index = 1 To TotalCount
                If ranks(index) = rank Then
                    PPTK_CommandAt index, commandName, shortcut, token, aliases
                    PPTK_AddResult commandName, shortcut, token, PPTK_CommandIcon(commandName)
                End If
            Next index
        Next rank
        lblSection.Caption = "RESULTS"
        lblTitle.Caption = ResultCount & " of " & TotalCount
    Else
        ReDim used(1 To TotalCount)
        picks = Split(PPTK_Suggestions(ContextKind), "|")
        For rank = LBound(picks) To UBound(picks)
            index = PPTK_CommandIndex(picks(rank))
            If index > 0 Then
                If Not used(index) Then
                    used(index) = True
                    PPTK_CommandAt index, commandName, shortcut, token, aliases
                    PPTK_AddResult commandName, shortcut, token, PPTK_CommandIcon(commandName)
                End If
            End If
        Next rank
        SuggestCount = ResultCount
        For index = 1 To TotalCount
            If Not used(index) Then
                PPTK_CommandAt index, commandName, shortcut, token, aliases
                PPTK_AddResult commandName, shortcut, token, PPTK_CommandIcon(commandName)
            End If
        Next index
        lblSection.Caption = "SUGGESTED " & ChrW$(&HB7) & " " & UCase$(PPTK_ContextLabel(ContextKind))
        lblTitle.Caption = TotalCount & " commands"
    End If

    If ResultCount > 0 Then SelectedIndex = 1 Else SelectedIndex = 0
    FirstVisible = 1
    fraResults.Visible = (ResultCount > 0)
    lblEmpty.Visible = (ResultCount = 0)
    PPTK_RenderRows
    PPTK_SyncScroll
    Painting = False
End Sub

Private Sub PPTK_AddResult(ByVal commandName As String, ByVal shortcut As String, ByVal token As String, ByVal icon As String)
    ResultCount = ResultCount + 1
    ResultNames(ResultCount) = commandName
    ResultKeys(ResultCount) = shortcut
    ResultTokens(ResultCount) = token
    ResultIcons(ResultCount) = icon
End Sub

Private Function PPTK_CommandIndex(ByVal wanted As String) As Long
    Dim index As Long
    Dim commandName As String
    Dim shortcut As String
    Dim token As String
    Dim aliases As String
    For index = 1 To TotalCount
        PPTK_CommandAt index, commandName, shortcut, token, aliases
        If commandName = wanted Then
            PPTK_CommandIndex = index
            Exit Function
        End If
    Next index
End Function

' Current PowerPoint context, read once. Nothing here changes the selection.
'   none   - nothing selected (or slides, or another view)
'   text   - a text cursor or text selection in a shape
'   table  - a table, its cells, or text inside it
'   group  - one group (or shapes inside one group)
'   one    - one shape
'   multi  - two or more AutoShapes, text boxes, or placeholders
'   mixed  - two or more objects that include pictures, tables, charts, groups...
Private Function PPTK_DetectContext() As String
    Dim rng As Object
    Dim index As Long
    Dim shapeKind As Long
    Dim plain As Boolean
    PPTK_DetectContext = "none"
    On Error GoTo Done
    If ActiveWindow Is Nothing Then Exit Function
    Select Case ActiveWindow.Selection.Type
        Case ppSelectionText
            If ActiveWindow.Selection.ShapeRange(1).HasTable Then
                PPTK_DetectContext = "table"
            Else
                PPTK_DetectContext = "text"
            End If
        Case ppSelectionShapes
            If ActiveWindow.Selection.HasChildShapeRange Then
                Set rng = ActiveWindow.Selection.ChildShapeRange
                If rng.Count = 1 Then
                    PPTK_DetectContext = "one"
                    Exit Function
                End If
            Else
                Set rng = ActiveWindow.Selection.ShapeRange
            End If
            If rng.Count = 1 Then
                If rng(1).HasTable Then
                    PPTK_DetectContext = "table"
                ElseIf rng(1).Type = msoGroup Then
                    PPTK_DetectContext = "group"
                Else
                    PPTK_DetectContext = "one"
                End If
            ElseIf rng.Count > 1 Then
                plain = True
                For index = 1 To rng.Count
                    shapeKind = rng(index).Type
                    If shapeKind <> msoAutoShape And shapeKind <> msoTextBox And shapeKind <> msoPlaceholder Then plain = False
                    If rng(index).HasTable Then plain = False
                Next index
                If plain Then PPTK_DetectContext = "multi" Else PPTK_DetectContext = "mixed"
            End If
    End Select
Done:
End Function

Private Function PPTK_ContextLabel(ByVal kind As String) As String
    Select Case kind
        Case "text": PPTK_ContextLabel = "Text editing"
        Case "table": PPTK_ContextLabel = "Table"
        Case "group": PPTK_ContextLabel = "Group"
        Case "one": PPTK_ContextLabel = "One shape"
        Case "multi": PPTK_ContextLabel = "Several shapes"
        Case "mixed": PPTK_ContextLabel = "Mixed selection"
        Case Else: PPTK_ContextLabel = "No selection"
    End Select
End Function

' Deterministic suggestions per context: command names from the registry,
' only commands that run in that context.
Private Function PPTK_Suggestions(ByVal kind As String) As String
    Select Case kind
        Case "text"
            PPTK_Suggestions = "Remove Margins|Paste Content Only|Increase Font Size|Decrease Font Size"
        Case "table"
            PPTK_Suggestions = "Add Right Column - Keep Bounds|Add Bottom Row - Keep Bounds|Delete Right Column - Keep Bounds|Delete Bottom Row - Keep Bounds|Remove Margins"
        Case "group"
            PPTK_Suggestions = "Remove Margins|Copy Formatting|Add Right TK|Add Bottom TK|Bring Forward"
        Case "one"
            PPTK_Suggestions = "Add Right TK|Add Bottom TK|Copy Formatting|Remove Margins|Bring Forward"
        Case "multi"
            PPTK_Suggestions = "Align Left|Align Top|Distribute Horizontally|Same Size|Unify Shape Type"
        Case "mixed"
            PPTK_Suggestions = "Align Left|Align Top|Distribute Horizontally|Same Width|Swap Shapes"
        Case Else
            PPTK_Suggestions = "Insert Textbox|Insert Rectangle|Insert New Slide|Insert Icon|Fix Slide Colors"
    End Select
End Function

Private Function PPTK_Normalize(ByVal value As String) As String
    value = LCase$(Trim$(value))
    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop
    PPTK_Normalize = value
End Function

' Returns 0-4 for a match, 99 for no match.
Private Function PPTK_Rank(ByVal query As String, ByVal commandName As String, ByVal shortcut As String, ByVal aliases As String) As Long
    Dim nameValue As String
    Dim keyValue As String
    Dim words() As String
    Dim index As Long

    If Len(query) = 0 Then
        PPTK_Rank = 0
        Exit Function
    End If
    PPTK_Rank = 99
    nameValue = PPTK_Normalize(commandName)
    keyValue = PPTK_Normalize(shortcut)
    If nameValue = query Or (Len(keyValue) > 0 And keyValue = query) Then
        PPTK_Rank = 0
    ElseIf Left$(nameValue, Len(query)) = query Then
        PPTK_Rank = 1
    ElseIf InStr(1, nameValue, query, vbBinaryCompare) > 0 Then
        PPTK_Rank = 2
    Else
        words = Split(PPTK_Normalize(aliases), " ")
        For index = LBound(words) To UBound(words)
            If Len(words(index)) > 0 And Left$(words(index), Len(query)) = query Then
                PPTK_Rank = 3
                Exit Function
            End If
        Next index
        If Len(keyValue) > 0 And InStr(1, keyValue, query, vbBinaryCompare) > 0 Then PPTK_Rank = 4
    End If
End Function

' Icon per command name. The pictures are hidden Image controls named
' ico_<key> (normal row) and ico_<key>_sel (selected row), stored in this
' form and built from src\PPTK\Brand\smartbar-icons\<key>.bmp and
' <key>-sel.bmp by Build-PptkSmartBarIcons.ps1. Opaque BMPs are used because
' a picture stored in a form loses GIF transparency and ICO colors.
Private Function PPTK_CommandIcon(ByVal commandName As String) As String
    Dim key As String
    Select Case commandName
        Case "Smart Bar": key = "pptk_smart_bar"
        Case "Fix Slide Colors": key = "pptk_fix_colors"
        Case "Find & Replace Colors": key = "pptk_replace_colors"
        Case "Align Process": key = "pptk_process"
        Case "Align Matrix": key = "pptk_matrix"
        Case "Align Central": key = "pptk_central"
        Case "Add Right Column - Keep Bounds": key = "sb_col_add_keep"
        Case "Add Bottom Row - Keep Bounds": key = "sb_row_add_keep"
        Case "Delete Right Column - Keep Bounds": key = "sb_col_del_keep"
        Case "Delete Bottom Row - Keep Bounds": key = "sb_row_del_keep"
        Case "Add Right TK": key = "pptk_tk_right"
        Case "Add Bottom TK": key = "pptk_tk_bottom"
        Case "Add Top TK": key = "pptk_tk_top"
        Case "Add Left TK": key = "pptk_tk_left"
        Case "Insert Textbox": key = "pptk_text"
        Case "Insert Rectangle": key = "pptk_shape_rect"
        Case "Insert Rounded Rectangle": key = "pptk_shape_round"
        Case "Insert Oval": key = "pptk_shape_oval"
        Case "Insert Chevron": key = "pptk_shape_chevron"
        Case "Insert Pentagon": key = "pptk_shape_pentagon"
        Case "Insert Triangle": key = "pptk_shape_triangle"
        Case "Insert Parallelogram": key = "pptk_shape_para"
        Case "Insert New Slide": key = "pptk_new_slide"
        Case "Insert Icon": key = "pptk_insert_icon"
        Case "Enclose Icons": key = "pptk_enclose"
        Case "Align Left": key = "pptk_align_left"
        Case "Align Right": key = "pptk_align_right"
        Case "Align Top": key = "pptk_align_top"
        Case "Align Bottom": key = "pptk_align_bottom"
        Case "Align Middle": key = "pptk_align_middle"
        Case "Align Center": key = "pptk_align_center"
        Case "Match Corner Roundness": key = "sb_corner"
        Case "Same Width": key = "pptk_same_width"
        Case "Same Height": key = "pptk_same_height"
        Case "Same Size": key = "pptk_same_size"
        Case "Distribute Horizontally": key = "pptk_dist_h"
        Case "Distribute Vertically": key = "pptk_dist_v"
        Case "Swap Shapes": key = "pptk_swap"
        Case "Swap Text": key = "pptk_swap_text"
        Case "Merge Textboxes": key = "pptk_merge_text"
        Case "Save Selected Slides": key = "pptk_save_slides"
        Case "Increase Horizontal Spacing": key = "pptk_gap_h_inc"
        Case "Decrease Horizontal Spacing": key = "pptk_gap_h_dec"
        Case "Increase Vertical Spacing": key = "pptk_gap_v_inc"
        Case "Decrease Vertical Spacing": key = "pptk_gap_v_dec"
        Case "Copy Formatting": key = "pptk_copy_format"
        Case "Paste Formatting": key = "pptk_paste_format"
        Case "Paste Content Only": key = "pptk_paste_text"
        Case "Remove Margins": key = "pptk_remove_margins"
        Case "Increase Font Size": key = "sb_font_up"
        Case "Decrease Font Size": key = "sb_font_down"
        Case "Bring Forward": key = "sb_bring_forward"
        Case "Send Backward": key = "sb_send_backward"
        Case "Add Right Column - Expand Bounds": key = "sb_col_add_expand"
        Case "Add Bottom Row - Expand Bounds": key = "sb_row_add_expand"
        Case "Delete Right Column - Shrink Bounds": key = "sb_col_del_shrink"
        Case "Delete Bottom Row - Shrink Bounds": key = "sb_row_del_shrink"
        Case "Flip Current Slide": key = "sb_flip_slide"
        Case "Flip Entire Presentation": key = "sb_flip_deck"
        Case "Unify Shape Type": key = "sb_unify"
        Case Else: key = ""
    End Select
    PPTK_CommandIcon = key
End Function

Private Function PPTK_IconPicture(ByVal key As String, ByVal isSelected As Boolean) As Object
    On Error Resume Next
    Set PPTK_IconPicture = Nothing
    If Len(key) = 0 Then Exit Function
    If isSelected Then
        Set PPTK_IconPicture = Me.Controls("ico_" & key & "_sel").Picture
    Else
        Set PPTK_IconPicture = Me.Controls("ico_" & key).Picture
    End If
End Function

Private Sub PPTK_RenderRows()
    Dim slot As Long
    Dim resultIndex As Long
    VisibleCount = 0
    For slot = 1 To PoolCount
        resultIndex = FirstVisible + slot - 1
        If resultIndex <= ResultCount Then
            VisibleCount = VisibleCount + 1
            RowName(slot).Caption = ResultNames(resultIndex)
            RowKey(slot).Caption = ResultKeys(resultIndex)
            Set RowIcon(slot).Picture = PPTK_IconPicture(ResultIcons(resultIndex), resultIndex = SelectedIndex)
            PPTK_PaintRow slot, (resultIndex = SelectedIndex)
            RowBg(slot).Visible = True
            RowIcon(slot).Visible = (Len(ResultIcons(resultIndex)) > 0)
            RowIcon(slot).ZOrder 0
            RowName(slot).Visible = True
            RowKey(slot).Visible = (Len(ResultKeys(resultIndex)) > 0)
        Else
            RowBg(slot).Visible = False
            RowIcon(slot).Visible = False
            RowName(slot).Visible = False
            RowKey(slot).Visible = False
        End If
    Next slot
    If RowDivider Is Nothing Then Exit Sub
    slot = SuggestCount + 1 - FirstVisible + 1
    If SuggestCount > 0 And SuggestCount < ResultCount And slot >= 2 And slot <= VisibleCount Then
        RowDivider.Left = 6
        RowDivider.Top = (slot - 1) * RowH
        RowDivider.Width = fraResults.Width - 12
        RowDivider.Height = 0.75
        RowDivider.Visible = True
        RowDivider.ZOrder 0
    Else
        RowDivider.Visible = False
    End If
End Sub

Private Sub PPTK_Reflow()
    Dim index As Long
    Dim rowWidth As Single
    Dim nameLeft As Single
    Dim keyLeft As Single
    Dim rowTop As Single

    If PoolCount < 1 Then Exit Sub
    rowWidth = fraResults.Width
    nameLeft = 6 + IconW + 8
    keyLeft = rowWidth - 8 - KeyW
    For index = 1 To PoolCount
        rowTop = (index - 1) * RowH
        RowBg(index).Left = 0
        RowBg(index).Top = rowTop
        RowBg(index).Width = rowWidth
        RowBg(index).Height = RowH
        RowIcon(index).Left = 6
        RowIcon(index).Top = rowTop + 7
        RowIcon(index).Width = 16
        RowIcon(index).Height = 16
        RowName(index).Left = nameLeft
        RowName(index).Top = rowTop + 6
        RowName(index).Width = keyLeft - nameLeft - 4
        RowName(index).Height = 18
        RowKey(index).Left = keyLeft
        RowKey(index).Top = rowTop + 8
        RowKey(index).Width = KeyW
        RowKey(index).Height = 15
    Next index
    On Error Resume Next
    fraResults.ScrollTop = 0
    fraResults.ScrollHeight = fraResults.Height
    fraResults.ScrollWidth = rowWidth
    On Error GoTo 0
End Sub

Private Sub PPTK_SelectRow(ByVal rowIndex As Long, ByVal scrollIntoView As Boolean)
    If rowIndex < 1 Or rowIndex > ResultCount Then Exit Sub
    SelectedIndex = rowIndex
    If scrollIntoView Then PPTK_EnsureVisible
    Painting = True
    PPTK_RenderRows
    PPTK_SyncScroll
    Painting = False
End Sub

Private Sub PPTK_PaintRow(ByVal index As Long, ByVal isSelected As Boolean)
    Dim fill As Long
    If index < 1 Or index > PoolCount Then Exit Sub
    If isSelected Then fill = ColorSelected Else fill = ColorPage
    RowBg(index).BackColor = fill
    RowName(index).BackColor = fill
    RowKey(index).BackColor = fill
End Sub

Private Sub PPTK_EnsureVisible()
    If SelectedIndex < 1 Then Exit Sub
    If SelectedIndex < FirstVisible Then
        FirstVisible = SelectedIndex
    ElseIf SelectedIndex > FirstVisible + VisibleRows - 1 Then
        FirstVisible = SelectedIndex - VisibleRows + 1
    End If
    If FirstVisible < 1 Then FirstVisible = 1
End Sub

Private Function PPTK_LastTop() As Long
    PPTK_LastTop = ResultCount - VisibleRows + 1
    If PPTK_LastTop < 1 Then PPTK_LastTop = 1
End Function

' Sizes and places the thumb for the current first row. The scrollbar is
' hidden when every result fits.
Private Sub PPTK_SyncScroll()
    Dim travel As Single
    Dim lastTop As Long
    If ScrollTrack Is Nothing Then Exit Sub
    lastTop = PPTK_LastTop()
    If FirstVisible > lastTop Then FirstVisible = lastTop
    ScrollTrack.Visible = (ResultCount > VisibleRows)
    ScrollThumb.Visible = ScrollTrack.Visible
    If Not ScrollTrack.Visible Then Exit Sub
    ScrollThumb.Height = ScrollTrack.Height * VisibleRows / ResultCount
    If ScrollThumb.Height < 18 Then ScrollThumb.Height = 18
    travel = ScrollTrack.Height - ScrollThumb.Height
    If lastTop > 1 Then
        ScrollThumb.Top = ScrollTrack.Top + travel * (FirstVisible - 1) / (lastTop - 1)
    Else
        ScrollThumb.Top = ScrollTrack.Top
    End If
End Sub

' Scrolls the list without changing the selected command.
Private Sub PPTK_ScrollTo(ByVal firstRow As Long)
    firstRow = PPTK_Clamp(firstRow, 1, PPTK_LastTop())
    If firstRow = FirstVisible Then Exit Sub
    FirstVisible = firstRow
    Painting = True
    PPTK_RenderRows
    PPTK_SyncScroll
    Painting = False
End Sub

Private Sub PPTK_RunSelected()
    Dim token As String
    If Busy Then Exit Sub
    If SelectedIndex < 1 Or SelectedIndex > ResultCount Then Exit Sub
    token = ResultTokens(SelectedIndex)
    Busy = True
    Me.Hide
    If Len(token) = 0 Then
        Busy = False
        Exit Sub
    End If
    If token = "CTRL K" Or token = "CTRL SPACE" Then
        Busy = False
        Me.Show vbModeless
        Exit Sub
    End If
    PPTK_RestoreSelection
    On Error GoTo Done
    PPTK_ExecuteToken token
Done:
    Busy = False
End Sub

Private Sub PPTK_ExecuteToken(ByVal token As String)
    Dim trimmed As String
    Dim procName As String
    trimmed = Trim$(token)
    If UCase$(Left$(trimmed, 5)) = "PROC " Then
        procName = Trim$(Mid$(trimmed, 6))
        If Len(procName) = 0 Then Exit Sub
        Application.Run procName
    Else
        PPTK_DispatchShortcut trimmed
    End If
End Sub

' The selection is kept as shape references, not names: names can repeat on a
' slide, and shapes selected inside a group must stay the group's children.
Private Sub PPTK_CaptureSelection()
    Dim index As Long
    Dim range As ShapeRange
    SelectedCount = 0
    On Error GoTo Clean
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    If ActiveWindow.Selection.HasChildShapeRange Then
        Set range = ActiveWindow.Selection.ChildShapeRange
    Else
        Set range = ActiveWindow.Selection.ShapeRange
    End If
    If range.Count < 1 Then Exit Sub
    ReDim SelectedShapes(1 To range.Count)
    For index = 1 To range.Count
        Set SelectedShapes(index) = range(index)
    Next index
    SelectedCount = range.Count
Clean:
End Sub

Private Sub PPTK_RestoreSelection()
    Dim index As Long
    If SelectedCount < 1 Then Exit Sub
    On Error GoTo Clean
    SelectedShapes(1).Select msoTrue
    For index = 2 To SelectedCount
        SelectedShapes(index).Select msoFalse
    Next index
Clean:
End Sub
