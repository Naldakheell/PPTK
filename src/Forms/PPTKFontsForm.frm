VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTKFontsForm
   Caption         =   "Replace Fonts"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PPTKFontsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' PPTK Replace Fonts window: Source font, Replacement font, Apply to (All text / Arabic only / English only),
' Replace and Close. The work is done by ModulePPTKReplaceFonts.PPTK_ReplaceFontsRun. Controls are created when
' the form loads (the designer holds none), so the form source is code only.
' Dark PPTK look: each font field is a text box with an arrow; its list is a PPTKDarkList (Smart Bar
' scrollbar) that drops down under it. "Font to replace" lists the fonts used in the presentation and is
' read-only; "Replace with" lists the installed fonts (PPTK_InstalledFonts), filters as you type, and
' accepts any typed name. Keys in a field: Down / Up open and move, Enter picks, Esc closes the list
' (Esc again closes the window). The buttons are PPTKFlatButton faces over the real buttons, which keep
' Enter (Replace) and Esc (Close).

Private Const ColorInk As Long = PPTK_DARK_INK
Private Const ColorMuted As Long = PPTK_DARK_MUTED
Private Const ColorError As Long = PPTK_DARK_ERROR
Private Const Pad As Single = 14
Private Const FieldWidth As Single = 250
Private Const ArrowW As Single = 18
Private Const ListRows As Long = 8
Private Const RowH As Single = 17

Private WithEvents cmdReplace As MSForms.CommandButton
Private WithEvents cmdClose As MSForms.CommandButton
Private WithEvents flatReplace As PPTKFlatButton
Private WithEvents flatClose As PPTKFlatButton
Private WithEvents txtSource As MSForms.TextBox
Private WithEvents txtTarget As MSForms.TextBox
Private WithEvents lblSourceArrow As MSForms.Label
Private WithEvents lblTargetArrow As MSForms.Label
Private WithEvents lstSource As PPTKDarkList
Private WithEvents lstTarget As PPTKDarkList
Private optAll As MSForms.OptionButton
Private optArabic As MSForms.OptionButton
Private optEnglish As MSForms.OptionButton
Private lblResult As MSForms.Label
Private Built As Boolean
Private NeedFill As Boolean
Private Filling As Boolean
Private TargetFonts() As String

' The window shows first; the font lists fill right after (UserForm_Activate), with a short "Loading fonts"
' note, so opening never looks frozen. The installed-font list is read once per PowerPoint session
' (PPTK_InstalledFonts caches it); the fonts used in the presentation are read on every open.
Public Sub PPTK_Open()
    PPTK_Build
    NeedFill = True
    lblResult.Caption = ""
    lstSource.Visible = False
    lstTarget.Visible = False
    Me.Show vbModal
End Sub

Private Sub UserForm_Activate()
    If Not NeedFill Then Exit Sub
    NeedFill = False
    PPTK_Say "Loading fonts...", False
    Me.MousePointer = fmMousePointerHourGlass
    Me.Repaint
    DoEvents
    PPTK_FillLists
    Me.MousePointer = fmMousePointerDefault
    PPTK_Say "", False
End Sub

Private Sub PPTK_Build()
    Dim y As Single
    If Built Then Exit Sub
    PPTK_StyleForm Me
    Me.Caption = "Replace Fonts"
    y = 12
    PPTK_Caption "lblSource", "Font to replace", y
    Set txtSource = PPTK_Field("txtSource", y + 16, True)
    Set lblSourceArrow = PPTK_Arrow("lblSourceArrow", txtSource)
    y = y + 50
    PPTK_Caption "lblTarget", "Replace with", y
    Set txtTarget = PPTK_Field("txtTarget", y + 16, False)
    Set lblTargetArrow = PPTK_Arrow("lblTargetArrow", txtTarget)
    y = y + 50
    PPTK_Caption "lblScope", "Apply to", y
    Set optAll = PPTK_Option("optAll", "All text", Pad, y + 16, 70)
    Set optArabic = PPTK_Option("optArabic", "Arabic only", Pad + 76, y + 16, 84)
    Set optEnglish = PPTK_Option("optEnglish", "English only", Pad + 166, y + 16, 90)
    optAll.Value = True
    y = y + 44
    Set lblResult = Me.Controls.Add("Forms.Label.1", "lblResult", True)
    lblResult.Left = Pad: lblResult.Top = y: lblResult.Width = FieldWidth: lblResult.Height = 42
    lblResult.Font.Name = "Segoe UI": lblResult.Font.Size = 8: lblResult.ForeColor = ColorMuted: lblResult.WordWrap = True
    lblResult.BackStyle = fmBackStyleTransparent
    y = y + 48
    Set cmdClose = Me.Controls.Add("Forms.CommandButton.1", "cmdClose", True)
    cmdClose.Caption = "Close": cmdClose.Cancel = True
    cmdClose.Width = 76: cmdClose.Height = 26: cmdClose.Top = y: cmdClose.Left = Pad + FieldWidth - 76
    Set cmdReplace = Me.Controls.Add("Forms.CommandButton.1", "cmdReplace", True)
    cmdReplace.Caption = "Replace": cmdReplace.Default = True
    cmdReplace.Width = 86: cmdReplace.Height = 26: cmdReplace.Top = y: cmdReplace.Left = cmdClose.Left - 92
    Set flatClose = New PPTKFlatButton
    flatClose.Attach Me, cmdClose, False
    Set flatReplace = New PPTKFlatButton
    flatReplace.Attach Me, cmdReplace, True
    ' The lists drop down over the controls below their field.
    Set lstSource = New PPTKDarkList
    lstSource.Init Me, "lstSource", Pad, txtSource.Top + txtSource.Height, FieldWidth, ListRows * RowH + 8, RowH
    lstSource.EmptyMessage = "No fonts in this presentation"
    Set lstTarget = New PPTKDarkList
    lstTarget.Init Me, "lstTarget", Pad, txtTarget.Top + txtTarget.Height, FieldWidth, ListRows * RowH + 8, RowH
    lstTarget.EmptyMessage = "No matching fonts (the typed name is used)"
    lstSource.Visible = False
    lstTarget.Visible = False
    Me.Width = FieldWidth + Pad * 2 + 16
    Me.Height = y + 26 + 44
    Built = True
End Sub

Private Sub PPTK_Caption(ByVal name As String, ByVal caption As String, ByVal y As Single)
    Dim label As MSForms.Label
    Set label = Me.Controls.Add("Forms.Label.1", name, True)
    label.Caption = caption: label.Left = Pad: label.Top = y: label.Width = FieldWidth: label.Height = 14
    label.Font.Name = "Segoe UI": label.Font.Size = 9: label.Font.Bold = True: label.ForeColor = ColorInk
    label.BackStyle = fmBackStyleTransparent
End Sub

Private Function PPTK_Field(ByVal name As String, ByVal y As Single, ByVal readOnly As Boolean) As MSForms.TextBox
    Dim box As MSForms.TextBox
    Set box = Me.Controls.Add("Forms.TextBox.1", name, True)
    box.Left = Pad: box.Top = y: box.Width = FieldWidth: box.Height = 20
    box.Font.Name = "Segoe UI": box.Font.Size = 9
    PPTK_StyleField box
    box.Locked = readOnly
    box.EnterKeyBehavior = False
    box.EnterFieldBehavior = fmEnterFieldBehaviorRecallSelection
    Set PPTK_Field = box
End Function

Private Function PPTK_Arrow(ByVal name As String, ByVal box As MSForms.TextBox) As MSForms.Label
    Dim arrow As MSForms.Label
    Set arrow = Me.Controls.Add("Forms.Label.1", name, True)
    arrow.Caption = ChrW$(&H25BE)
    arrow.Font.Name = "Segoe UI": arrow.Font.Size = 9
    arrow.TextAlign = fmTextAlignCenter
    arrow.ForeColor = ColorMuted
    arrow.BackStyle = fmBackStyleTransparent
    arrow.Left = box.Left + box.Width - ArrowW
    arrow.Top = box.Top + 3
    arrow.Width = ArrowW - 2
    arrow.Height = box.Height - 5
    arrow.ZOrder 0
    Set PPTK_Arrow = arrow
End Function

Private Function PPTK_Option(ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal w As Single) As MSForms.OptionButton
    Dim opt As MSForms.OptionButton
    Set opt = Me.Controls.Add("Forms.OptionButton.1", name, True)
    opt.Caption = caption: opt.GroupName = "PPTKFontScope"
    opt.Left = x: opt.Top = y: opt.Width = w: opt.Height = 18
    opt.Font.Name = "Segoe UI": opt.Font.Size = 9: opt.ForeColor = ColorInk
    opt.BackStyle = fmBackStyleTransparent
    Set PPTK_Option = opt
End Function

' "Font to replace": the fonts used in the presentation. "Replace with": the installed fonts (or, when they
' cannot be read, the used fonts and a few common ones).
Private Sub PPTK_FillLists()
    Dim used() As String
    Dim installed As String
    Dim keepSource As String
    keepSource = txtSource.Text
    used = Split(PPTK_FontsInUse(), vbLf)
    If Len(Join(used, "")) = 0 Then
        lstSource.SetItems Empty
    Else
        lstSource.SetItems used
    End If
    installed = PPTK_InstalledFonts()
    If Len(installed) = 0 Then installed = PPTK_FontsInUse() & vbLf & "Arial" & vbLf & "Tahoma" & vbLf & "Calibri" & vbLf & "Segoe UI" & vbLf & "Aptos" & vbLf & "Dubai" & vbLf & "Sakkal Majalla" & vbLf & "Traditional Arabic" & vbLf & "Simplified Arabic"
    TargetFonts = Split(installed, vbLf)
    PPTK_FilterTarget ""
    Filling = True
    If lstSource.IndexOf(keepSource) > 0 Then txtSource.Text = keepSource Else txtSource.Text = ""
    Filling = False
End Sub

Private Sub PPTK_FilterTarget(ByVal query As String)
    Dim matches() As String
    Dim count As Long
    Dim index As Long
    query = Trim$(query)
    ReDim matches(0 To UBound(TargetFonts) - LBound(TargetFonts))
    count = 0
    For index = LBound(TargetFonts) To UBound(TargetFonts)
        If Len(TargetFonts(index)) > 0 Then
            If Len(query) = 0 Or InStr(1, TargetFonts(index), query, vbTextCompare) > 0 Then
                matches(count) = TargetFonts(index)
                count = count + 1
            End If
        End If
    Next index
    If count = 0 Then
        lstTarget.SetItems Empty
    Else
        ReDim Preserve matches(0 To count - 1)
        lstTarget.SetItems matches
    End If
End Sub

' ---- dropdown behaviour

Private Sub PPTK_ShowList(ByVal list As PPTKDarkList, ByVal box As MSForms.TextBox)
    If list Is lstSource Then lstTarget.Visible = False Else lstSource.Visible = False
    If list.ListIndex = 0 And Len(box.Text) > 0 Then
        If list.IndexOf(box.Text) > 0 Then list.ListIndex = list.IndexOf(box.Text)
    End If
    list.Visible = True
End Sub

Private Sub PPTK_HideLists()
    lstSource.Visible = False
    lstTarget.Visible = False
End Sub

Private Sub PPTK_Pick(ByVal list As PPTKDarkList, ByVal box As MSForms.TextBox)
    If list.ListIndex < 1 Then Exit Sub
    Filling = True
    box.Text = list.ItemText(list.ListIndex)
    Filling = False
    list.Visible = False
    On Error Resume Next
    box.SetFocus
    box.SelStart = Len(box.Text)
End Sub

' Keys in a font field. Returns True when the key was used (so Enter/Esc do not reach the buttons).
Private Function PPTK_FieldKey(ByVal list As PPTKDarkList, ByVal box As MSForms.TextBox, ByVal keyCode As Long) As Boolean
    Select Case keyCode
        Case vbKeyDown, vbKeyUp, vbKeyPageDown, vbKeyPageUp
            If Not list.Visible Then PPTK_ShowList list, box
            list.HandleKey keyCode
            PPTK_FieldKey = True
        Case vbKeyReturn
            If list.Visible And list.ListIndex > 0 Then
                PPTK_Pick list, box
                PPTK_FieldKey = True
            ElseIf list.Visible Then
                list.Visible = False
                PPTK_FieldKey = True
            End If
        Case vbKeyEscape
            If list.Visible Then
                list.Visible = False
                PPTK_FieldKey = True
            End If
    End Select
End Function

Private Sub txtSource_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PPTK_FieldKey(lstSource, txtSource, KeyCode) Then KeyCode = 0
End Sub

Private Sub txtTarget_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If PPTK_FieldKey(lstTarget, txtTarget, KeyCode) Then KeyCode = 0
End Sub

Private Sub txtSource_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lstSource.Visible Then lstSource.Visible = False Else PPTK_ShowList lstSource, txtSource
End Sub

Private Sub lblSourceArrow_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    txtSource.SetFocus
    On Error GoTo 0
    If lstSource.Visible Then lstSource.Visible = False Else PPTK_ShowList lstSource, txtSource
End Sub

Private Sub lblTargetArrow_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    txtTarget.SetFocus
    On Error GoTo 0
    If lstTarget.Visible Then lstTarget.Visible = False Else PPTK_ShowList lstTarget, txtTarget
End Sub

Private Sub txtTarget_Change()
    If Filling Then Exit Sub
    PPTK_FilterTarget txtTarget.Text
    If lstTarget.Count > 0 Then lstTarget.ListIndex = 1
    PPTK_ShowList lstTarget, txtTarget
End Sub

Private Sub lstSource_Picked()
    PPTK_Pick lstSource, txtSource
End Sub

Private Sub lstTarget_Picked()
    PPTK_Pick lstTarget, txtTarget
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PPTK_HideLists
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    flatReplace.Leave
    flatClose.Leave
    lstSource.Leave
    lstTarget.Leave
End Sub

' ---- buttons

Private Sub flatReplace_Click()
    cmdReplace_Click
End Sub

Private Sub flatClose_Click()
    cmdClose_Click
End Sub

Private Sub cmdReplace_Click()
    Dim scope As String
    Dim spec As String
    Dim result As String
    PPTK_HideLists
    If Len(Trim$(txtSource.Text)) = 0 Or Len(Trim$(txtTarget.Text)) = 0 Then
        PPTK_Say "Choose the font to replace and the new font.", True
        Exit Sub
    End If
    If optArabic.Value Then
        scope = "Arabic"
    ElseIf optEnglish.Value Then
        scope = "English"
    Else
        scope = "All"
    End If
    spec = Trim$(txtSource.Text) & "|" & Trim$(txtTarget.Text) & "|" & scope
    Me.MousePointer = fmMousePointerHourGlass
    result = PPTK_ReplaceFontsRun(spec)
    If result = "nobackup" Then
        Me.MousePointer = fmMousePointerDefault
        If MsgBox("A backup copy could not be saved." & vbCr & "Replace anyway? You can still undo with Ctrl+Z.", vbYesNo + vbExclamation + vbDefaultButton2, "PPTK") <> vbYes Then Exit Sub
        Me.MousePointer = fmMousePointerHourGlass
        result = PPTK_ReplaceFontsRun(spec & "|force")
    End If
    Me.MousePointer = fmMousePointerDefault
    If Left$(result, 6) = "error=" Then
        PPTK_Say Mid$(result, 7) & ".", True
        Exit Sub
    End If
    PPTK_Say PPTK_Summary(result), False
    PPTK_FillLists
End Sub

Private Function PPTK_Summary(ByVal result As String) As String
    Dim text As String
    text = "Done: " & PPTK_Field2(result, "runs") & " text change(s), " & PPTK_Field2(result, "theme") & " theme font(s)."
    If Val(PPTK_Field2(result, "left")) > 0 Then text = text & " " & PPTK_Field2(result, "left") & " place(s) still use it."
    If Val(PPTK_Field2(result, "charts")) > 0 Then text = text & " Charts were not checked (" & PPTK_Field2(result, "charts") & ")."
    If Val(PPTK_Field2(result, "failed")) > 0 Then text = text & " " & PPTK_Field2(result, "failed") & " could not be changed."
    If Len(PPTK_Field2(result, "backup")) > 0 Then text = text & " Backup saved in Documents\PPTK Font Backups."
    PPTK_Summary = text
End Function

Private Function PPTK_Field2(ByVal result As String, ByVal name As String) As String
    Dim parts() As String
    Dim index As Long
    parts = Split(result, ";")
    For index = LBound(parts) To UBound(parts)
        If Left$(parts(index), Len(name) + 1) = name & "=" Then PPTK_Field2 = Mid$(parts(index), Len(name) + 2): Exit Function
    Next index
End Function

Private Sub PPTK_Say(ByVal message As String, ByVal isError As Boolean)
    lblResult.Caption = message
    If isError Then lblResult.ForeColor = ColorError Else lblResult.ForeColor = ColorMuted
End Sub

Private Sub cmdClose_Click()
    PPTK_HideLists
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    Me.Hide
End Sub
