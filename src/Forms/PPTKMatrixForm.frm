VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTKMatrixForm
   Caption         =   "Align Matrix"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   StartUpPosition =   1  'CenterOwner
   BackColor       =   &00FBF9F7&
   Begin MSForms.Label lblAccent
      BackColor       =   &00ECBA23&
      Height          =   80
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
   Begin MSForms.Label lblGuide
      Caption         =   "Move across the squares, then click."
      Height          =   320
      Left            =   280
      TabIndex        =   1
      Top             =   240
      Width           =   4240
   End
   Begin MSForms.Label lblSize
      Caption         =   "Choose a size"
      Height          =   320
      Left            =   280
      TabIndex        =   2
      Top             =   3120
      Width           =   4240
   End
   Begin MSForms.Label lblError
      Caption         =   ""
      ForeColor       =   &002C26A4&
      Height          =   280
      Left            =   280
      TabIndex        =   3
      Top             =   3480
      Width           =   4240
   End
   Begin MSForms.CommandButton cmdMore
      Caption         =   "Sizes up to 12"
      Height          =   480
      Left            =   280
      TabIndex        =   4
      Top             =   3880
      Width           =   1960
   End
   Begin MSForms.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   480
      Left            =   3080
      TabIndex        =   5
      Top             =   3880
      Width           =   1440
   End
End
Attribute VB_Name = "PPTKMatrixForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Grid picker for Align Matrix. The layout math stays in PPTK_AlignMatrix.
' Squares are created when the form loads and kept by this saved form code.

Private Const ColorPage As Long = 16513527
Private Const ColorCard As Long = 16777215
Private Const ColorInk As Long = 3879984
Private Const ColorMuted As Long = &H706B6B
Private Const ColorAccent As Long = 15514147
Private Const ColorWash As Long = 16576978
Private Const ColorLine As Long = 15591651
Private Const ColorDim As Long = &HF3F2F2
Private Const ColorDimLine As Long = &HE0E0E0
Private Const ColorBad As Long = &HE8E4E4
Private Const ColorError As Long = &H2C26A4
Private Const CellSize As Single = 18
Private Const CellGap As Single = 3
Private Const Pad As Single = 14
Private Const MaxExtent As Long = 12
Private Const FirstExtent As Long = 6

Private Slots(1 To 12, 1 To 12) As PPTKMatrixCell
Private Built As Boolean
Private Accepted As Boolean
Private ShapeCount As Long
Private VisibleExtent As Long
Private HoverRow As Long
Private HoverCol As Long
Private ChosenRows As Long
Private ChosenCols As Long

Public Function PPTK_Ask(ByVal count As Long, ByRef rows As Long, ByRef cols As Long) As Boolean
    ShapeCount = count
    If ShapeCount < 1 Then ShapeCount = 1
    Accepted = False
    ChosenRows = 0
    ChosenCols = 0
    HoverRow = rows
    HoverCol = cols
    If HoverRow < 1 Then HoverRow = 1
    If HoverCol < 1 Then HoverCol = 1
    If HoverRow > MaxExtent Then HoverRow = MaxExtent
    If HoverCol > MaxExtent Then HoverCol = MaxExtent
    If rows > FirstExtent Or cols > FirstExtent Or ShapeCount > FirstExtent * FirstExtent Then
        VisibleExtent = MaxExtent
    Else
        VisibleExtent = FirstExtent
    End If
    PPTK_BuildGrid
    PPTK_Layout
    PPTK_Paint
    Me.Show vbModal
    PPTK_Ask = Accepted
    If Accepted Then
        rows = ChosenRows
        cols = ChosenCols
    End If
End Function

Public Sub PPTK_Hover(ByVal row As Long, ByVal col As Long)
    If row < 1 Or col < 1 Then Exit Sub
    If row > VisibleExtent Or col > VisibleExtent Then Exit Sub
    If row = HoverRow And col = HoverCol Then Exit Sub
    HoverRow = row
    HoverCol = col
    PPTK_Paint
End Sub

Public Sub PPTK_Choose(ByVal row As Long, ByVal col As Long)
    If Accepted Then Exit Sub
    If row < 1 Or col < 1 Then Exit Sub
    If row > VisibleExtent Or col > VisibleExtent Then Exit Sub
    If row > MaxExtent Or col > MaxExtent Then Exit Sub
    If row * col < ShapeCount Then Exit Sub
    ChosenRows = row
    ChosenCols = col
    Accepted = True
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    PPTK_StyleChrome
    VisibleExtent = FirstExtent
    PPTK_BuildGrid
End Sub

Private Sub UserForm_Activate()
    PPTK_Layout
    PPTK_Paint
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    Accepted = False
    Me.Hide
End Sub

Private Sub cmdMore_Click()
    VisibleExtent = MaxExtent
    PPTK_Layout
    PPTK_Paint
End Sub

Private Sub cmdCancel_Click()
    Accepted = False
    Me.Hide
End Sub

Private Sub PPTK_BuildGrid()
    Dim row As Long
    Dim col As Long
    Dim name As String
    Dim box As MSForms.Label
    If Built Then Exit Sub
    For row = 1 To MaxExtent
        For col = 1 To MaxExtent
            name = "PPTKCell_" & Format$(row, "00") & "_" & Format$(col, "00")
            Set box = Nothing
            On Error Resume Next
            Set box = Me.Controls(name)
            On Error GoTo 0
            If box Is Nothing Then
                Set box = Me.Controls.Add("Forms.Label.1", name, True)
            End If
            box.Caption = ""
            box.Width = CellSize
            box.Height = CellSize
            box.BorderStyle = fmBorderStyleSingle
            box.SpecialEffect = fmSpecialEffectFlat
            box.BackStyle = fmBackStyleOpaque
            Set Slots(row, col) = New PPTKMatrixCell
            Slots(row, col).Attach Me, box, row, col
        Next col
    Next row
    Built = True
End Sub

Private Sub PPTK_StyleChrome()
    PPTK_StyleForm Me
    Me.BackColor = ColorPage
    Me.Font.Name = "Segoe UI"
    lblAccent.Caption = ""
    lblAccent.BackColor = ColorAccent
    lblAccent.Left = 0
    lblAccent.Top = 0
    lblAccent.Height = 4
    PPTK_Font lblGuide, 9, False, ColorMuted
    lblGuide.BackStyle = fmBackStyleTransparent
    lblGuide.Caption = "Hover a square, then click."
    PPTK_Font lblSize, 10, True, ColorInk
    lblSize.BackStyle = fmBackStyleTransparent
    lblSize.TextAlign = fmTextAlignCenter
    PPTK_Font lblError, 8, False, ColorError
    lblError.BackStyle = fmBackStyleTransparent
    lblError.TextAlign = fmTextAlignCenter
    PPTK_Button cmdCancel, False
    PPTK_Button cmdMore, False
    cmdCancel.Caption = "Cancel"
    cmdMore.Caption = "Sizes up to 12"
End Sub

Private Sub PPTK_Layout()
    Dim row As Long
    Dim col As Long
    Dim grid As Single
    Dim inner As Single
    Dim origin As Single
    Dim topGrid As Single
    grid = VisibleExtent * CellSize + (VisibleExtent - 1) * CellGap
    inner = grid + (Pad * 2)
    If inner < 268 Then inner = 268
    origin = (inner - grid) / 2
    topGrid = 32
    lblAccent.Width = inner
    lblGuide.Left = Pad
    lblGuide.Top = 10
    lblGuide.Width = inner - (Pad * 2)
    lblGuide.Height = 16
    For row = 1 To MaxExtent
        For col = 1 To MaxExtent
            With Slots(row, col).Box
                .Visible = (row <= VisibleExtent And col <= VisibleExtent)
                .Left = origin + (col - 1) * (CellSize + CellGap)
                .Top = topGrid + (row - 1) * (CellSize + CellGap)
                .Width = CellSize
                .Height = CellSize
            End With
        Next col
    Next row
    lblSize.Left = Pad
    lblSize.Top = topGrid + grid + 10
    lblSize.Width = inner - (Pad * 2)
    lblSize.Height = 18
    lblError.Left = Pad
    lblError.Top = lblSize.Top + 18
    lblError.Width = lblSize.Width
    lblError.Height = 14
    cmdCancel.Top = lblError.Top + 20
    cmdMore.Top = cmdCancel.Top
    cmdCancel.Height = 26
    cmdMore.Height = 26
    cmdCancel.Width = 76
    cmdMore.Width = 112
    cmdCancel.Left = inner - Pad - cmdCancel.Width
    cmdMore.Left = Pad
    cmdMore.Visible = (VisibleExtent < MaxExtent)
    Me.Width = inner + 16
    Me.Height = cmdCancel.Top + cmdCancel.Height + 44
End Sub

Private Sub PPTK_Paint()
    Dim row As Long
    Dim col As Long
    Dim inside As Boolean
    Dim valid As Boolean
    Dim mark As String
    valid = (HoverRow >= 1 And HoverCol >= 1 And HoverRow * HoverCol >= ShapeCount)
    For row = 1 To VisibleExtent
        For col = 1 To VisibleExtent
            inside = (row <= HoverRow And col <= HoverCol)
            If inside And valid Then
                Slots(row, col).Paint ColorWash, ColorAccent
            ElseIf inside Then
                Slots(row, col).Paint ColorBad, ColorDimLine
            ElseIf row * col < ShapeCount Then
                Slots(row, col).Paint ColorDim, ColorDimLine
            Else
                Slots(row, col).Paint ColorCard, ColorLine
            End If
        Next col
    Next row
    If HoverRow < 1 Or HoverCol < 1 Then
        lblSize.Caption = "Choose a size"
        lblError.Caption = ""
        Exit Sub
    End If
    mark = HoverCol & " columns " & ChrW$(215) & " " & HoverRow & " rows"
    lblSize.Caption = mark
    If valid Then
        lblError.Caption = ""
    Else
        lblError.Caption = "Needs at least " & ShapeCount & " cells."
    End If
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
