VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTKSmartBarForm
   Caption         =   "Smart Bar"
   ClientHeight    =   5240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   StartUpPosition =   1  'CenterOwner
   BackColor       =   &00FBF9F7&
   Begin MSForms.Label lblAccent
      BackColor       =   &00ECBA23&
      Height          =   80
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6720
   End
   Begin MSForms.TextBox txtFilter
      BorderColor     =   &00ECBA23&
      BorderStyle     =   1
      Height          =   520
      Left            =   280
      SpecialEffect   =   0
      TabIndex        =   1
      Top             =   280
      Width           =   6160
   End
   Begin MSForms.Label lblCommand
      Caption         =   "Command"
      Height          =   280
      Left            =   320
      TabIndex        =   2
      Top             =   920
      Width           =   3600
   End
   Begin MSForms.Label lblShortcut
      Caption         =   "Shortcut"
      Height          =   280
      Left            =   4600
      TabIndex        =   3
      Top             =   920
      Width           =   1800
   End
   Begin MSForms.ListBox lstResults
      Height          =   3360
      Left            =   280
      TabIndex        =   4
      Top             =   1240
      Width           =   6160
   End
   Begin MSForms.Label lblEmpty
      Caption         =   "No matching commands"
      Height          =   360
      Left            =   280
      TabIndex        =   5
      Top             =   2480
      Visible         =   0   'False
      Width           =   6160
   End
   Begin MSForms.Label lblHint
      Caption         =   "Up and Down select. Enter runs. Esc closes."
      Height          =   320
      Left            =   280
      TabIndex        =   6
      Top             =   4760
      Width           =   6160
   End
End
Attribute VB_Name = "PPTKSmartBarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Compact command palette. The list is the same 52-command registry.
' MSForms cannot tint the list selection, so that highlight stays the Windows color.

Private Const ColorPage As Long = 16513527
Private Const ColorCard As Long = 16777215
Private Const ColorInk As Long = 3879984
Private Const ColorMuted As Long = &H706B6B
Private Const ColorAccent As Long = 15514147
Private Const ColorLine As Long = 15591651
Private Const Pad As Single = 14

Private ResultTokens() As String
Private SelectedNames() As String
Private SelectedSlide As Long
Private SelectedCount As Long
Private Busy As Boolean

Private Sub UserForm_Initialize()
    PPTK_Style
    txtFilter.Text = ""
    PPTK_Refill
    PPTK_CaptureSelection
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    txtFilter.SetFocus
    PPTK_CaptureSelection
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    Me.Hide
End Sub

Private Sub txtFilter_Change()
    PPTK_Refill
End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    PPTK_RunSelected
End Sub

Private Sub txtFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    PPTK_HandleKey KeyCode
End Sub

Private Sub lstResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    PPTK_HandleKey KeyCode
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    PPTK_HandleKey KeyCode
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
        If lstResults.ListIndex < lstResults.ListCount - 1 Then
            lstResults.ListIndex = lstResults.ListIndex + 1
        End If
    ElseIf KeyCode = vbKeyUp Then
        KeyCode = 0
        If lstResults.ListIndex > 0 Then
            lstResults.ListIndex = lstResults.ListIndex - 1
        End If
    End If
End Sub

Private Sub PPTK_Style()
    Dim inner As Single
    PPTK_StyleForm Me
    Me.BackColor = ColorPage
    Me.Font.Name = "Segoe UI"
    inner = 336
    lblAccent.Caption = ""
    lblAccent.BackColor = ColorAccent
    lblAccent.Left = 0
    lblAccent.Top = 0
    lblAccent.Height = 4
    lblAccent.Width = inner

    PPTK_StyleField txtFilter
    PPTK_Font txtFilter, 11, False, ColorInk
    txtFilter.Left = Pad
    txtFilter.Top = 12
    txtFilter.Width = inner - (Pad * 2)
    txtFilter.Height = 26
    txtFilter.BackColor = ColorCard
    txtFilter.BorderStyle = fmBorderStyleSingle
    txtFilter.BorderColor = ColorAccent
    txtFilter.SpecialEffect = fmSpecialEffectFlat
    txtFilter.MultiLine = False
    txtFilter.EnterKeyBehavior = False

    PPTK_Font lblCommand, 8, False, ColorMuted
    lblCommand.Caption = "Command"
    lblCommand.Left = Pad + 2
    lblCommand.Top = 44
    lblCommand.Width = 180
    lblCommand.Height = 14
    lblCommand.BackStyle = fmBackStyleTransparent

    PPTK_Font lblShortcut, 8, False, ColorMuted
    lblShortcut.Caption = "Shortcut"
    lblShortcut.Left = Pad + 214
    lblShortcut.Top = 44
    lblShortcut.Width = 96
    lblShortcut.Height = 14
    lblShortcut.BackStyle = fmBackStyleTransparent

    lstResults.Left = Pad
    lstResults.Top = 60
    lstResults.Width = inner - (Pad * 2)
    lstResults.Height = 168
    lstResults.ColumnCount = 2
    lstResults.ColumnWidths = "198 pt;90 pt"
    lstResults.IntegralHeight = True
    lstResults.BorderStyle = fmBorderStyleSingle
    lstResults.BorderColor = ColorLine
    lstResults.BackColor = ColorCard
    lstResults.ForeColor = ColorInk
    lstResults.SpecialEffect = fmSpecialEffectFlat
    lstResults.Font.Name = "Segoe UI"
    lstResults.Font.Size = 9
    lstResults.MultiSelect = fmMultiSelectSingle

    PPTK_Font lblEmpty, 9, False, ColorMuted
    lblEmpty.Caption = "No matching commands"
    lblEmpty.TextAlign = fmTextAlignCenter
    lblEmpty.BackColor = ColorCard
    lblEmpty.BorderStyle = fmBorderStyleSingle
    lblEmpty.BorderColor = ColorLine
    lblEmpty.Left = lstResults.Left
    lblEmpty.Top = lstResults.Top
    lblEmpty.Width = lstResults.Width
    lblEmpty.Height = lstResults.Height
    lblEmpty.Visible = False

    PPTK_Font lblHint, 8, False, ColorMuted
    lblHint.Caption = "Up and Down select. Enter runs. Esc closes."
    lblHint.BackStyle = fmBackStyleTransparent
    lblHint.Left = Pad
    lblHint.Width = inner - (Pad * 2)
    lblHint.Height = 14
    lblHint.Top = lstResults.Top + lstResults.Height + 8
    Me.Width = inner + 16
    Me.Height = lblHint.Top + lblHint.Height + 28
End Sub

Private Sub PPTK_Font(ByVal target As Object, ByVal size As Single, ByVal bold As Boolean, ByVal color As Long)
    target.Font.Name = "Segoe UI"
    target.Font.Size = size
    target.Font.Bold = bold
    target.ForeColor = color
End Sub

Private Sub PPTK_Refill()
    Dim index As Long
    Dim commandName As String
    Dim shortcut As String
    Dim token As String
    Dim aliases As String
    Dim query As String
    Dim found As Long

    query = LCase$(Trim$(txtFilter.Text))
    lstResults.Clear
    found = 0
    ReDim ResultTokens(1 To PPTK_CommandCount)
    For index = 1 To PPTK_CommandCount
        PPTK_CommandAt index, commandName, shortcut, token, aliases
        If PPTK_Matches(query, commandName, shortcut, aliases) Then
            found = found + 1
            ResultTokens(found) = token
            lstResults.AddItem commandName
            lstResults.List(lstResults.ListCount - 1, 1) = shortcut
        End If
    Next index
    If found > 0 Then
        lstResults.Visible = True
        lblEmpty.Visible = False
        lstResults.ListIndex = 0
    Else
        lstResults.Visible = False
        lblEmpty.Visible = True
    End If
End Sub

Private Function PPTK_Matches(ByVal query As String, ByVal commandName As String, ByVal shortcut As String, ByVal aliases As String) As Boolean
    If Len(query) = 0 Then
        PPTK_Matches = True
    Else
        PPTK_Matches = InStr(1, LCase$(commandName), query, vbTextCompare) > 0 _
            Or InStr(1, LCase$(shortcut), query, vbTextCompare) > 0 _
            Or InStr(1, LCase$(aliases), query, vbTextCompare) > 0
    End If
End Function

Private Sub PPTK_RunSelected()
    Dim token As String
    If Busy Then Exit Sub
    If lstResults.ListIndex < 0 Then Exit Sub
    Busy = True
    token = ResultTokens(lstResults.ListIndex + 1)
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
    PPTK_DispatchShortcut token
Done:
    Busy = False
End Sub

Private Sub PPTK_CaptureSelection()
    Dim index As Long
    SelectedCount = 0
    On Error GoTo Clean
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    SelectedCount = ActiveWindow.Selection.ShapeRange.Count
    If SelectedCount < 1 Then Exit Sub
    ReDim SelectedNames(1 To SelectedCount)
    For index = 1 To SelectedCount
        SelectedNames(index) = ActiveWindow.Selection.ShapeRange(index).Name
    Next index
    SelectedSlide = ActiveWindow.View.Slide.SlideIndex
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
