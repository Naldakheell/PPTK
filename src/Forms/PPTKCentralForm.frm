VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTKCentralForm
   Caption         =   "Align Central"
   ClientHeight    =   2960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   StartUpPosition =   1  'CenterOwner
   BackColor       =   &00FBF9F7&
   Begin MSForms.Label lblAccent
      BackColor       =   &00ECBA23&
      Height          =   80
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5040
   End
   Begin MSForms.Label lblHint
      Caption         =   "Distance from the center, in points. Shapes keep their size."
      Height          =   560
      Left            =   320
      TabIndex        =   1
      Top             =   280
      Width           =   4400
   End
   Begin MSForms.Label lblRadius
      Caption         =   "Radius"
      Height          =   360
      Left            =   320
      TabIndex        =   2
      Top             =   1080
      Width           =   1040
   End
   Begin MSForms.TextBox txtRadius
      BorderColor     =   &00ECBA23&
      BorderStyle     =   1
      Height          =   440
      Left            =   1480
      SpecialEffect   =   0
      TabIndex        =   3
      Top             =   1040
      Width           =   1280
   End
   Begin MSForms.Label lblError
      Caption         =   ""
      ForeColor       =   &002C26A4&
      Height          =   320
      Left            =   320
      TabIndex        =   4
      Top             =   1600
      Width           =   4400
   End
   Begin MSForms.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   480
      Left            =   1680
      TabIndex        =   5
      Top             =   2200
      Width           =   1440
   End
   Begin MSForms.CommandButton cmdOk
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   480
      Left            =   3280
      TabIndex        =   6
      Top             =   2200
      Width           =   1440
   End
End
Attribute VB_Name = "PPTKCentralForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Radius dialog. Placement, angles, and the default radius stay in PPTK_AlignCentral.

Private Const ColorPage As Long = 16513527
Private Const ColorCard As Long = 16777215
Private Const ColorInk As Long = 3879984
Private Const ColorMuted As Long = &H706B6B
Private Const ColorAccent As Long = 15514147
Private Const ColorError As Long = &H2C26A4

Private Accepted As Boolean
Private Chosen As Single

Public Function PPTK_Ask(ByRef radius As Single) As Boolean
    Accepted = False
    Chosen = radius
    lblError.Caption = ""
    txtRadius.Text = CStr(CLng(radius))
    Me.Show vbModal
    PPTK_Ask = Accepted
    If Accepted Then radius = Chosen
End Function

Private Sub UserForm_Initialize()
    PPTK_Style
End Sub

Private Sub UserForm_Activate()
    PPTK_Style
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    Accepted = False
    Me.Hide
End Sub

Private Sub txtRadius_Change()
    lblError.Caption = ""
End Sub

Private Sub cmdOk_Click()
    Dim value As Double
    On Error GoTo Bad
    value = CDbl(txtRadius.Text)
    If value < 0 Or value > 2000 Then GoTo Bad
    Chosen = CSng(value)
    Accepted = True
    Me.Hide
    Exit Sub
Bad:
    lblError.Caption = "Enter a radius from 0 to 2000 points."
End Sub

Private Sub cmdCancel_Click()
    Accepted = False
    Me.Hide
End Sub

Private Sub PPTK_Style()
    PPTK_StyleForm Me
    Me.BackColor = ColorPage
    Me.Font.Name = "Segoe UI"
    lblAccent.Caption = ""
    lblAccent.BackColor = ColorAccent
    lblAccent.Left = 0
    lblAccent.Top = 0
    lblAccent.Height = 4
    lblAccent.Width = 280

    PPTK_Font lblHint, 9, False, ColorMuted
    lblHint.BackStyle = fmBackStyleTransparent
    lblHint.Caption = "Distance from the center, in points. Shapes keep their size."
    lblHint.Left = 16
    lblHint.Top = 14
    lblHint.Width = 248
    lblHint.Height = 32

    PPTK_Font lblRadius, 10, False, ColorInk
    lblRadius.BackStyle = fmBackStyleTransparent
    lblRadius.Caption = "Radius"
    lblRadius.Left = 16
    lblRadius.Top = 54
    lblRadius.Width = 56
    lblRadius.Height = 20

    PPTK_StyleField txtRadius
    PPTK_Font txtRadius, 11, False, ColorInk
    txtRadius.Left = 76
    txtRadius.Top = 50
    txtRadius.Width = 72
    txtRadius.Height = 24
    txtRadius.BackColor = ColorCard
    txtRadius.BorderStyle = fmBorderStyleSingle
    txtRadius.BorderColor = ColorAccent
    txtRadius.SpecialEffect = fmSpecialEffectFlat

    PPTK_Font lblError, 8, False, ColorError
    lblError.BackStyle = fmBackStyleTransparent
    lblError.Left = 16
    lblError.Top = 82
    lblError.Width = 248
    lblError.Height = 16

    PPTK_Button cmdCancel, False
    PPTK_Button cmdOk, True
    cmdOk.Caption = "Apply"
    cmdCancel.Caption = "Cancel"
    cmdOk.Top = 108
    cmdCancel.Top = 108
    cmdOk.Height = 26
    cmdCancel.Height = 26
    cmdOk.Width = 76
    cmdCancel.Width = 76
    cmdOk.Left = 280 - 16 - cmdOk.Width
    cmdCancel.Left = cmdOk.Left - 8 - cmdCancel.Width
    Me.Width = 420
    Me.Height = cmdOk.Top + cmdOk.Height + 28
End Sub

Private Sub PPTK_Font(ByVal target As Object, ByVal size As Single, ByVal bold As Boolean, ByVal color As Long)
    target.Font.Name = "Segoe UI"
    target.Font.Size = size
    target.Font.Bold = bold
    target.ForeColor = color
End Sub

Private Sub PPTK_Button(ByVal button As MSForms.CommandButton, ByVal primary As Boolean)
    button.Font.Name = "Segoe UI"
    button.Font.Size = 9
    button.Font.Bold = primary
    If primary Then
        PPTK_StylePrimaryButton button
    Else
        PPTK_StyleSecondaryButton button
    End If
End Sub
