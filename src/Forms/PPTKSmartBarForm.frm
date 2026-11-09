VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTKSmartBarForm 
   BorderStyle     =   1  'Single
   Caption         =   "PPTK Smart Bar"
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

' Compact dark command palette.
' Rows are labels inside fraResults. txtFilter keeps the live search.
' A row runs the token already stored by ModulePPTKSmartBar.
' Shortcut tokens go to PPTK_DispatchShortcut. PROC tokens call that public procedure.
' Sizes are VBA points. Results are virtualized; the native scrollbar stays disabled.

Private Const ColorPage As Long = 2103312
Private Const ColorField As Long = 3549980
Private Const ColorAccent As Long = 15514147
Private Const ColorSelected As Long = 5522212
Private Const ColorInk As Long = 16513268
Private Const ColorMuted As Long = 13352100
Private Const ClientW As Single = 440
Private Const ClientH As Single = 390
Private Const Pad As Single = 18
Private Const RowH As Single = 46
Private Const VisibleRows As Long = 4
Private Const KeyCol As Single = 76
Private Const IconCol As Single = 28

Private RowHooks As Collection
Private RowBg() As MSForms.Label
Private RowBar() As MSForms.Label
Private RowName() As MSForms.Label
Private RowKey() As MSForms.Label
Private RowIcon() As MSForms.Label
Private ResultNames() As String
Private ResultKeys() As String
Private ResultTokens() As String
Private ResultIcons() As String
Private SelectedNames() As String
Private PoolCount As Long
Private VisibleCount As Long
Private ResultCount As Long
Private SelectedIndex As Long
Private FirstVisible As Long
Private SelectedSlide As Long
Private SelectedCount As Long
Private Busy As Boolean
Private Painting As Boolean

Private Sub UserForm_Initialize()
    Set RowHooks = New Collection
    PPTK_ApplyChrome
    PPTK_BuildPool
    txtFilter.text = ""
    PPTK_Refill
    PPTK_CaptureSelection
End Sub

Private Sub UserForm_Activate()
    PPTK_Layout
    On Error Resume Next
    txtFilter.SetFocus
    On Error GoTo 0
    PPTK_CaptureSelection
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    Me.Hide
End Sub

Private Sub txtFilter_Change()
    lblCue.Visible = (Len(txtFilter.text) = 0)
    PPTK_Refill
End Sub

Private Sub txtFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    PPTK_HandleKey KeyCode
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    PPTK_HandleKey KeyCode
End Sub

Private Sub fraResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    PPTK_HandleKey KeyCode
End Sub

Private Sub lblCue_Click()
    On Error Resume Next
    txtFilter.SetFocus
End Sub

Private Sub lblField_Click()
    On Error Resume Next
    txtFilter.SetFocus
End Sub

Private Sub lblSearchEnter_Click()
    On Error Resume Next
    txtFilter.SetFocus
End Sub

Private Sub lblSearchIcon_Click()
    On Error Resume Next
    txtFilter.SetFocus
End Sub

Private Sub fraResults_Click()
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

Private Sub PPTK_HandleKey(ByVal KeyCode As MSForms.ReturnInteger)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Me.Hide
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        PPTK_RunSelected
        Exit Sub
    End If
    If KeyCode = vbKeyDown Then
        KeyCode = 0
        If SelectedIndex < ResultCount Then PPTK_SelectRow SelectedIndex + 1, True
    ElseIf KeyCode = vbKeyUp Then
        KeyCode = 0
        If SelectedIndex > 1 Then PPTK_SelectRow SelectedIndex - 1, True
    End If
End Sub

Private Sub PPTK_ApplyChrome()
    PPTK_StyleForm Me
    Me.Caption = "PPTK Smart Bar"
    Me.width = ClientW
    Me.height = ClientH
    Me.BackColor = ColorPage
    Me.ForeColor = ColorPage
    Me.Font.name = "Segoe UI"

    lblAccent.Visible = False

    PPTK_Font lblEsc, 9, False, ColorMuted
    lblEsc.Caption = "ESC"
    lblEsc.TextAlign = fmTextAlignRight
    lblEsc.BackStyle = fmBackStyleTransparent

    PPTK_Font lblTitle, 9, False, ColorMuted
    lblTitle.Caption = "Command palette"
    lblTitle.TextAlign = fmTextAlignLeft
    lblTitle.BackStyle = fmBackStyleTransparent

    PPTK_Font lblBrand, 14, True, ColorInk
    lblBrand.Caption = "PPTK"
    lblBrand.TextAlign = fmTextAlignLeft
    lblBrand.BackStyle = fmBackStyleTransparent

    PPTK_Font lblBadge, 18, True, ColorPage
    lblBadge.Caption = ChrW$(&H2318)
    lblBadge.TextAlign = fmTextAlignCenter
    lblBadge.BackStyle = fmBackStyleOpaque
    lblBadge.BackColor = ColorAccent
    lblBadge.BorderStyle = fmBorderStyleNone

    lblField.Caption = ""
    lblField.BackColor = ColorField
    lblField.BorderStyle = fmBorderStyleSingle
    lblField.BorderColor = ColorSelected
    lblField.SpecialEffect = fmSpecialEffectFlat

    PPTK_Font lblCue, 10, False, ColorMuted
    lblCue.Caption = "Search commands..."
    lblCue.TextAlign = fmTextAlignLeft
    lblCue.BackColor = ColorField
    lblCue.BackStyle = fmBackStyleOpaque
    lblCue.BorderStyle = fmBorderStyleNone

    PPTK_Font lblSearchEnter, 14, False, ColorMuted
    lblSearchEnter.Caption = ChrW$(&H21B5)
    lblSearchEnter.TextAlign = fmTextAlignCenter
    lblSearchEnter.BackColor = ColorField
    lblSearchEnter.BackStyle = fmBackStyleOpaque

    PPTK_Font lblSearchIcon, 14, False, ColorAccent
    lblSearchIcon.Caption = ChrW$(&H2315)
    lblSearchIcon.TextAlign = fmTextAlignCenter
    lblSearchIcon.BackColor = ColorField
    lblSearchIcon.BackStyle = fmBackStyleOpaque

    txtFilter.BackColor = ColorField
    txtFilter.ForeColor = ColorInk
    txtFilter.BorderStyle = fmBorderStyleNone
    txtFilter.SpecialEffect = fmSpecialEffectFlat
    txtFilter.Font.name = "Segoe UI"
    txtFilter.Font.size = 12
    txtFilter.Font.bold = False
    txtFilter.Multiline = False
    txtFilter.EnterKeyBehavior = False
    txtFilter.TextAlign = fmTextAlignLeft

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
    lblEmpty.Caption = "No matching commands"
    lblEmpty.TextAlign = fmTextAlignCenter
    lblEmpty.BackStyle = fmBackStyleTransparent
    lblEmpty.BorderStyle = fmBorderStyleNone
    lblEmpty.Visible = False

    PPTK_Font lblSection, 8, True, ColorMuted
    lblSection.Caption = "COMMANDS"
    lblSection.TextAlign = fmTextAlignLeft
    lblSection.BackStyle = fmBackStyleTransparent

    lblFooterLine.Caption = ""
    lblFooterLine.BackColor = ColorSelected
    lblFooterLine.BackStyle = fmBackStyleOpaque
    lblFooterLine.BorderStyle = fmBorderStyleNone

    PPTK_Font lblRun, 9, False, ColorMuted
    lblRun.Caption = "Run " & ChrW$(&H21B5)
    lblRun.BackStyle = fmBackStyleTransparent

    PPTK_Font lblNavigate, 9, False, ColorMuted
    lblNavigate.Caption = "Navigate " & ChrW$(&H2193) & " " & ChrW$(&H2191)
    lblNavigate.TextAlign = fmTextAlignRight
    lblNavigate.BackStyle = fmBackStyleTransparent

    lblHint.Caption = ""
    lblHint.Visible = False
    lblAccent.Caption = ""
    lblAccent.Visible = False
End Sub

Private Sub PPTK_Layout()
    Dim extraW As Single
    Dim extraH As Single
    Dim widthNow As Single
    Dim fieldTop As Single
    Dim frameTop As Single
    Dim inner As Single

    extraW = Me.width - Me.InsideWidth
    extraH = Me.height - Me.InsideHeight
    If extraW < 0 Or extraW > 48 Then extraW = 0
    If extraH < 0 Or extraH > 80 Then extraH = 0
    Me.width = ClientW + extraW
    Me.height = ClientH + extraH
    If Me.InsideWidth > 50 Then Me.width = Me.width + (ClientW - Me.InsideWidth)
    If Me.InsideHeight > 50 Then Me.height = Me.height + (ClientH - Me.InsideHeight)
    widthNow = Me.InsideWidth
    If widthNow < 80 Then widthNow = ClientW
    inner = widthNow - (Pad * 2)

    lblBadge.width = 38
    lblBadge.height = 34
    lblBadge.left = Pad
    lblBadge.Top = 10

    lblBrand.width = 48
    lblBrand.height = 24
    lblBrand.left = lblBadge.left + lblBadge.width + 10
    lblBrand.Top = 15

    lblTitle.width = 110
    lblTitle.height = 18
    lblTitle.left = lblBrand.left + lblBrand.width + 8
    lblTitle.Top = 20

    lblEsc.width = 42
    lblEsc.height = 18
    lblEsc.left = widthNow - Pad - lblEsc.width
    lblEsc.Top = 20

    fieldTop = 60
    lblField.left = Pad
    lblField.Top = fieldTop
    lblField.width = inner
    lblField.height = 48

    lblSearchIcon.left = lblField.left + 12
    lblSearchIcon.Top = fieldTop + 13
    lblSearchIcon.width = 24
    lblSearchIcon.height = 22

    lblSearchEnter.left = lblField.left + lblField.width - 36
    lblSearchEnter.Top = fieldTop + 13
    lblSearchEnter.width = 24
    lblSearchEnter.height = 22

    lblCue.left = lblField.left + 42
    lblCue.Top = fieldTop + 14
    lblCue.width = lblField.width - 84
    lblCue.height = 20

    txtFilter.left = lblField.left + 42
    txtFilter.Top = fieldTop + 8
    txtFilter.height = 30
    txtFilter.width = lblField.width - 84
    If txtFilter.width < 40 Then txtFilter.width = 40

    lblHint.left = 0
    lblHint.Top = 0
    lblHint.width = 0
    lblHint.height = 0
    lblAccent.left = 0
    lblAccent.Top = 0
    lblAccent.width = 0
    lblAccent.height = 0

    lblSection.left = Pad
    lblSection.Top = fieldTop + 61
    lblSection.width = inner
    lblSection.height = 16

    frameTop = fieldTop + 82
    fraResults.left = Pad
    fraResults.Top = frameTop
    fraResults.width = inner
    fraResults.height = RowH * VisibleRows
    fraResults.ScrollBars = fmScrollBarsNone
    fraResults.KeepScrollBarsVisible = fmScrollBarsNone

    lblEmpty.left = fraResults.left
    lblEmpty.Top = fraResults.Top + (fraResults.height - 22) / 2
    lblEmpty.width = fraResults.width
    lblEmpty.height = 22

    lblFooterLine.left = Pad
    lblFooterLine.Top = fraResults.Top + fraResults.height + 14
    lblFooterLine.width = inner
    lblFooterLine.height = 1

    lblRun.left = Pad
    lblRun.Top = lblFooterLine.Top + 14
    lblRun.width = 90
    lblRun.height = 18

    lblNavigate.left = widthNow - Pad - 120
    lblNavigate.Top = lblRun.Top
    lblNavigate.width = 120
    lblNavigate.height = 18

    lblField.ZOrder 1
    txtFilter.ZOrder 0
    lblCue.ZOrder 0
    lblSearchEnter.ZOrder 0
    lblSearchIcon.ZOrder 0
    lblEmpty.ZOrder 0
    PPTK_Reflow
End Sub

Private Sub PPTK_Font(ByVal target As Object, ByVal size As Single, ByVal bold As Boolean, ByVal color As Long)
    target.Font.name = "Segoe UI"
    target.Font.size = size
    target.Font.bold = bold
    target.ForeColor = color
End Sub

Private Sub PPTK_BuildPool()
    Dim needed As Long
    Dim index As Long
    needed = VisibleRows
    If PoolCount >= needed And Not RowHooks Is Nothing Then Exit Sub
    Set RowHooks = New Collection
    ReDim RowBg(1 To needed)
    ReDim RowBar(1 To needed)
    ReDim RowName(1 To needed)
    ReDim RowKey(1 To needed)
    ReDim RowIcon(1 To needed)
    For index = 1 To needed
        Set RowBg(index) = PPTK_MakeLabel("dyn_bg_" & Format$(index, "00"))
        Set RowBar(index) = PPTK_MakeLabel("dyn_bar_" & Format$(index, "00"))
        Set RowName(index) = PPTK_MakeLabel("dyn_name_" & Format$(index, "00"))
        Set RowKey(index) = PPTK_MakeLabel("dyn_key_" & Format$(index, "00"))
        Set RowIcon(index) = PPTK_MakeLabel("dyn_icon_" & Format$(index, "00"))
        PPTK_PrepareRow index
        PPTK_Bind RowBg(index), index
        PPTK_Bind RowBar(index), index
        PPTK_Bind RowName(index), index
        PPTK_Bind RowKey(index), index
        PPTK_Bind RowIcon(index), index
    Next index
    PoolCount = needed
End Sub

Private Function PPTK_MakeLabel(ByVal controlName As String) As MSForms.Label
    Dim existing As MSForms.control
    On Error Resume Next
    Set existing = fraResults.Controls(controlName)
    On Error GoTo 0
    If existing Is Nothing Then
        Set PPTK_MakeLabel = fraResults.Controls.Add("Forms.Label.1", controlName, True)
    Else
        Set PPTK_MakeLabel = existing
    End If
End Function

Private Sub PPTK_PrepareRow(ByVal index As Long)
    PPTK_PrepareLabel RowBg(index), False, False
    PPTK_PrepareLabel RowBar(index), False, False
    PPTK_PrepareLabel RowName(index), False, False
    PPTK_Font RowName(index), 11, False, ColorInk
    PPTK_PrepareLabel RowKey(index), True, False
    PPTK_Font RowKey(index), 9, False, ColorMuted
    RowKey(index).TextAlign = fmTextAlignCenter
    PPTK_PrepareLabel RowIcon(index), False, False
    PPTK_Font RowIcon(index), 12, False, ColorMuted
    RowIcon(index).TextAlign = fmTextAlignCenter
    RowBg(index).Visible = False
    RowBar(index).Visible = False
    RowName(index).Visible = False
    RowKey(index).Visible = False
    RowIcon(index).Visible = False
End Sub

Private Sub PPTK_PrepareLabel(ByVal labelControl As MSForms.Label, ByVal alignRight As Boolean, ByVal wrapText As Boolean)
    labelControl.Caption = ""
    labelControl.BorderStyle = fmBorderStyleNone
    labelControl.SpecialEffect = fmSpecialEffectFlat
    labelControl.BackStyle = fmBackStyleOpaque
    labelControl.BackColor = ColorPage
    labelControl.WordWrap = wrapText
    labelControl.AutoSize = False
    If alignRight Then
        labelControl.TextAlign = fmTextAlignRight
    Else
        labelControl.TextAlign = fmTextAlignLeft
    End If
End Sub

Private Sub PPTK_Bind(ByVal labelControl As MSForms.Label, ByVal index As Long)
    Dim hook As PPTKSmartBarRow
    Set hook = New PPTKSmartBarRow
    hook.Initialize labelControl, Me, index
    RowHooks.Add hook
End Sub

Private Sub PPTK_Refill()
    Dim index As Long
    Dim commandName As String
    Dim shortcut As String
    Dim token As String
    Dim aliases As String
    Dim query As String
    Dim found As Long

    If PoolCount < 1 Then PPTK_BuildPool
    Painting = True
    query = LCase$(Trim$(txtFilter.text))
    found = 0
    ReDim ResultNames(1 To PPTK_CommandCount)
    ReDim ResultKeys(1 To PPTK_CommandCount)
    ReDim ResultTokens(1 To PPTK_CommandCount)
    ReDim ResultIcons(1 To PPTK_CommandCount)
    For index = 1 To PPTK_CommandCount
        PPTK_CommandAt index, commandName, shortcut, token, aliases
        If PPTK_Matches(query, commandName, shortcut, aliases) Then
            found = found + 1
            ResultNames(found) = commandName
            ResultKeys(found) = shortcut
            ResultTokens(found) = token
            ResultIcons(found) = PPTK_CommandGlyph(commandName)
        End If
    Next index
    ResultCount = found
    If found > 0 Then
        SelectedIndex = 1
        FirstVisible = 1
    Else
        SelectedIndex = 0
        FirstVisible = 1
    End If
    fraResults.Visible = (found > 0)
    lblEmpty.Visible = (found = 0)
    PPTK_RenderRows
    Painting = False
End Sub

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
            RowIcon(slot).Caption = ResultIcons(resultIndex)
            RowBg(slot).Visible = True
            RowName(slot).Visible = True
            RowKey(slot).Visible = True
            RowIcon(slot).Visible = True
            RowBar(slot).Visible = False
            PPTK_PaintRow slot, (resultIndex = SelectedIndex)
        Else
            RowBg(slot).Visible = False
            RowBar(slot).Visible = False
            RowName(slot).Visible = False
            RowKey(slot).Visible = False
            RowIcon(slot).Visible = False
        End If
    Next slot
    PPTK_Reflow
    For slot = 1 To VisibleCount
        PPTK_FitKey slot
    Next slot
End Sub

Private Sub PPTK_FitKey(ByVal index As Long)
    Dim needed As Single
    Dim rowWidth As Single
    If index < 1 Or index > PoolCount Then Exit Sub
    If Len(RowKey(index).Caption) = 0 Then
        RowKey(index).width = 0
        Exit Sub
    End If
    RowKey(index).AutoSize = True
    needed = RowKey(index).width + 18
    RowKey(index).AutoSize = False
    If needed < 40 Then needed = 40
    If needed > KeyCol Then needed = KeyCol
    rowWidth = PPTK_ContentWidth()
    RowKey(index).width = needed
    RowKey(index).height = 22
    RowKey(index).left = rowWidth - 12 - needed
    RowKey(index).TextAlign = fmTextAlignCenter
End Sub

Private Function PPTK_CommandGlyph(ByVal commandName As String) As String
    Dim value As String
    value = LCase$(commandName)
    If InStr(value, "matrix") > 0 Then
        PPTK_CommandGlyph = ChrW$(&H25A6)
    ElseIf InStr(value, "column") > 0 Then
        PPTK_CommandGlyph = ChrW$(&H25A5)
    ElseIf InStr(value, "row") > 0 Then
        PPTK_CommandGlyph = ChrW$(&H25A4)
    ElseIf InStr(value, "size") > 0 Or InStr(value, "stretch") > 0 Then
        PPTK_CommandGlyph = ChrW$(&H2197)
    ElseIf InStr(value, "color") > 0 Then
        PPTK_CommandGlyph = ChrW$(&H25C9)
    Else
        PPTK_CommandGlyph = ChrW$(&H203A)
    End If
End Function

Private Function PPTK_Matches(ByVal query As String, ByVal commandName As String, ByVal shortcut As String, ByVal aliases As String) As Boolean
    If Len(query) = 0 Then
        PPTK_Matches = True
    Else
        PPTK_Matches = InStr(1, LCase$(commandName), query, vbTextCompare) > 0 _
            Or InStr(1, LCase$(shortcut), query, vbTextCompare) > 0 _
            Or InStr(1, LCase$(aliases), query, vbTextCompare) > 0
    End If
End Function

Private Sub PPTK_Reflow()
    Dim index As Long
    Dim rowWidth As Single
    Dim iconLeft As Single
    Dim keyLeft As Single
    Dim nameWidth As Single

    If PoolCount < 1 Then Exit Sub
    rowWidth = PPTK_ContentWidth()
    iconLeft = 12
    keyLeft = rowWidth - 12 - KeyCol
    nameWidth = keyLeft - (iconLeft + IconCol) - 16
    If nameWidth < 20 Then nameWidth = 20
    For index = 1 To PoolCount
        RowBg(index).left = 0
        RowBg(index).Top = (index - 1) * RowH
        RowBg(index).width = rowWidth
        RowBg(index).height = RowH
        RowBar(index).left = 0
        RowBar(index).Top = (index - 1) * RowH
        RowBar(index).width = 0
        RowBar(index).height = RowH
        RowIcon(index).left = iconLeft
        RowIcon(index).Top = (index - 1) * RowH + 12
        RowIcon(index).width = IconCol
        RowIcon(index).height = 22
        RowName(index).left = iconLeft + IconCol + 8
        RowName(index).Top = (index - 1) * RowH + 12
        RowName(index).width = nameWidth
        RowName(index).height = 22
        RowKey(index).left = keyLeft
        RowKey(index).Top = (index - 1) * RowH + 12
        RowKey(index).width = KeyCol
        RowKey(index).height = 22
        RowKey(index).TextAlign = fmTextAlignCenter
    Next index
    On Error Resume Next
    fraResults.ScrollTop = 0
    fraResults.ScrollLeft = 0
    fraResults.ScrollHeight = fraResults.height
    fraResults.ScrollWidth = rowWidth
    On Error GoTo 0
End Sub

Private Function PPTK_ContentWidth() As Single
    Dim inner As Single
    inner = fraResults.InsideWidth
    If inner < 40 Then inner = fraResults.width
    PPTK_ContentWidth = inner - 2
    If PPTK_ContentWidth < 40 Then PPTK_ContentWidth = 40
End Function

Private Sub PPTK_SelectRow(ByVal rowIndex As Long, ByVal scrollIntoView As Boolean)
    If rowIndex < 1 Or rowIndex > ResultCount Then Exit Sub
    SelectedIndex = rowIndex
    If scrollIntoView Then PPTK_EnsureVisible
    Painting = True
    PPTK_RenderRows
    Painting = False
End Sub

Private Sub PPTK_PaintRow(ByVal index As Long, ByVal selected As Boolean)
    Dim fill As Long
    If index < 1 Or index > PoolCount Then Exit Sub
    If selected Then fill = ColorSelected Else fill = ColorPage
    RowBg(index).BackColor = fill
    RowName(index).BackColor = fill
    RowIcon(index).BackColor = fill
    RowKey(index).BackColor = ColorField
    RowName(index).ForeColor = ColorInk
    RowKey(index).ForeColor = ColorMuted
    RowIcon(index).ForeColor = IIf(selected, ColorAccent, ColorMuted)
    RowBar(index).Visible = False
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

Private Sub PPTK_RunSelected()
    Dim token As String
    If Busy Then Exit Sub
    If SelectedIndex < 1 Or SelectedIndex > ResultCount Then Exit Sub
    Busy = True
    token = ResultTokens(SelectedIndex)
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
    If UCase$(left$(trimmed, 5)) = "PROC " Then
        procName = Trim$(Mid$(trimmed, 6))
        If Len(procName) = 0 Then Exit Sub
        Application.Run procName
    Else
        PPTK_DispatchShortcut trimmed
    End If
End Sub

Private Sub PPTK_CaptureSelection()
    Dim index As Long
    SelectedCount = 0
    On Error GoTo Clean
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    SelectedCount = ActiveWindow.Selection.ShapeRange.count
    If SelectedCount < 1 Then Exit Sub
    ReDim SelectedNames(1 To SelectedCount)
    For index = 1 To SelectedCount
        SelectedNames(index) = ActiveWindow.Selection.ShapeRange(index).name
    Next index
    SelectedSlide = ActiveWindow.view.Slide.SlideIndex
Clean:
End Sub

Private Sub PPTK_RestoreSelection()
    Dim index As Long
    Dim sld As Slide
    If SelectedCount < 1 Then Exit Sub
    On Error GoTo Clean
    Set sld = ActivePresentation.Slides(SelectedSlide)
    sld.Shapes(SelectedNames(1)).Select
    For index = 2 To SelectedCount
        sld.Shapes(SelectedNames(index)).Select False
    Next index
Clean:
End Sub

