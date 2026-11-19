VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPTKFlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A flat dark PPTK button (the Find & Replace Colors look) drawn over an MSForms CommandButton.
' The CommandButton stays underneath and keeps its job: Default (Enter), Cancel (Esc), Tab order and its
' Click code. The flat face is two labels (a box and a centred caption); a click on it raises Click,
' which the form forwards to the button's own Click code. Call Sync after the button moves, changes
' caption or is enabled/disabled; call Leave from the form's MouseMove to drop the hover colour.

Public Event Click()

Private WithEvents Box As MSForms.Label
Private WithEvents Face As MSForms.Label
Private Target As MSForms.CommandButton
Private IsPrimary As Boolean
Private Hot As Boolean

Public Sub Attach(ByVal host As Object, ByVal button As MSForms.CommandButton, ByVal primary As Boolean)
    Set Target = button
    IsPrimary = primary
    Set Box = host.Controls.Add("Forms.Label.1", button.Name & "_flatBox", True)
    Set Face = host.Controls.Add("Forms.Label.1", button.Name & "_flatText", True)
    Box.Caption = ""
    Box.BackStyle = fmBackStyleOpaque
    Box.BorderStyle = fmBorderStyleSingle
    Box.SpecialEffect = fmSpecialEffectFlat
    Face.BackStyle = fmBackStyleTransparent
    Face.BorderStyle = fmBorderStyleNone
    Face.TextAlign = fmTextAlignCenter
    Face.WordWrap = False
    Face.Font.Name = "Segoe UI"
    Face.Font.Size = 9
    Sync
End Sub

Public Sub Sync()
    If Target Is Nothing Then Exit Sub
    Box.Left = Target.Left
    Box.Top = Target.Top
    Box.Width = Target.Width
    Box.Height = Target.Height
    Face.Caption = Target.Caption
    Face.Left = Box.Left
    Face.Width = Box.Width
    Face.Height = 15
    Face.Top = Box.Top + (Box.Height - Face.Height) / 2
    Box.Visible = Target.Visible
    Face.Visible = Target.Visible
    PPTK_Paint
    Box.ZOrder 0
    Face.ZOrder 0
End Sub

Public Sub Leave()
    If Not Hot Then Exit Sub
    Hot = False
    PPTK_Paint
End Sub

Private Sub PPTK_Paint()
    If Not Target.Enabled Then
        Box.BackColor = PPTK_DARK_SURFACE
        Box.BorderColor = PPTK_DARK_LINE
        Face.ForeColor = PPTK_DARK_DIM
    ElseIf IsPrimary Then
        If Hot Then Box.BackColor = PPTK_PRIMARY_HOT Else Box.BackColor = PPTK_COLOR_PRIMARY
        Box.BorderColor = Box.BackColor
        Face.ForeColor = PPTK_DARK_PAGE
    Else
        If Hot Then Box.BackColor = PPTK_DARK_SELECTED Else Box.BackColor = PPTK_DARK_SURFACE
        If Hot Then Box.BorderColor = PPTK_DARK_STRONG_LINE Else Box.BorderColor = PPTK_DARK_LINE
        Face.ForeColor = PPTK_DARK_INK
    End If
End Sub

Private Sub PPTK_Hover()
    If Hot Or Not Target.Enabled Then Exit Sub
    Hot = True
    PPTK_Paint
End Sub

Private Sub PPTK_Fire()
    If Target Is Nothing Then Exit Sub
    If Target.Enabled And Target.Visible Then RaiseEvent Click
End Sub

Private Sub Box_Click()
    PPTK_Fire
End Sub

Private Sub Face_Click()
    PPTK_Fire
End Sub

Private Sub Box_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PPTK_Hover
End Sub

Private Sub Face_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PPTK_Hover
End Sub
