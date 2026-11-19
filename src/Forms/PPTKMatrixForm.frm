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

' Interactive grid picker for Align Matrix (Auxy style). The layout math stays in ModulePPTKSmartAlign.
' The column under the mouse sets the columns; the row sets the rows, raised to ceil(shapes / columns) when
' that is too few, so every square is a valid choice. Green squares are the shapes (filled row by row),
' white squares inside the layout are empty cells, the square under the mouse is blue.
' Moving over the grid previews the layout on the slide (PPTK_MatrixPreview); a click keeps it; Esc, Cancel
' or closing the form returns PPTK_Ask = False and the caller puts the shapes back.
' The grid starts at 6 x 6 (or larger for many shapes) and grows by one when the mouse reaches its last row
' or column, up to the shape count and at most MatrixMaxExtent (30). Squares are created as the grid grows.

' Dark PPTK window palette (ModulePPTKBrand, the Smart Bar's look). Cells: green = a shape, raised
' surface = empty cell inside the layout, dark surface = outside it, cyan = the cell under the mouse.
Private Const ColorPage As Long = PPTK_DARK_PAGE
Private Const ColorCard As Long = PPTK_DARK_SELECTED
Private Const ColorInk As Long = PPTK_DARK_INK
Private Const ColorMuted As Long = PPTK_DARK_MUTED
Private Const ColorLine As Long = PPTK_DARK_STRONG_LINE
Private Const ColorDim As Long = PPTK_DARK_SURFACE
Private Const ColorDimLine As Long = PPTK_DARK_LINE
Private Const ColorFilled As Long = &H96BE20
Private Const ColorFilledLine As Long = &H7A9A1A
Private Const ColorHover As Long = PPTK_COLOR_PRIMARY
Private Const ColorHoverLine As Long = PPTK_DARK_INK
Private Const CellSize As Single = 18
Private Const CellGap As Single = 3
Private Const Pad As Single = 14
Private Const StartExtent As Long = 6

Private Slots(1 To 30, 1 To 30) As PPTKMatrixCell
Private Accepted As Boolean
Private WithEvents flatCancel As PPTKFlatButton
Private ShapeCount As Long
Private VisRows As Long
Private VisCols As Long
Private LimitRows As Long
Private LimitCols As Long
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
    LimitRows = PPTK_Clamp(ShapeCount, StartExtent, MatrixMaxExtent)
    LimitCols = LimitRows
    HoverCol = PPTK_Clamp(cols, 1, LimitCols)
    HoverRow = PPTK_Clamp(PPTK_RowsFor(rows, HoverCol), 1, LimitRows)
    VisCols = PPTK_Clamp(HoverCol + 1, StartExtent, LimitCols)
    VisRows = PPTK_Clamp(HoverRow + 1, StartExtent, LimitRows)
    PPTK_EnsureCells
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
    Dim grown As Boolean
    If Accepted Then Exit Sub
    If row < 1 Or col < 1 Or row > VisRows Or col > VisCols Then Exit Sub
    If row = HoverRow And col = HoverCol Then Exit Sub
    HoverRow = row
    HoverCol = col
    If col = VisCols And VisCols < LimitCols Then VisCols = VisCols + 1: grown = True
    If row = VisRows And VisRows < LimitRows Then VisRows = VisRows + 1: grown = True
    If grown Then
        PPTK_EnsureCells
        PPTK_Layout
    End If
    PPTK_Paint
    PPTK_MatrixPreview PPTK_RowsFor(HoverRow, HoverCol) & "|" & HoverCol
    DoEvents
End Sub

Public Sub PPTK_Choose(ByVal row As Long, ByVal col As Long)
    If Accepted Then Exit Sub
    If row < 1 Or col < 1 Or row > VisRows Or col > VisCols Then Exit Sub
    ChosenCols = col
    ChosenRows = PPTK_RowsFor(row, col)
    Accepted = True
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    PPTK_StyleChrome
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

Private Sub cmdCancel_Click()
    Accepted = False
    Me.Hide
End Sub

Private Sub flatCancel_Click()
    cmdCancel_Click
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not flatCancel Is Nothing Then flatCancel.Leave
End Sub

' Rows for a layout: the row under the mouse, or enough rows for every shape.
Private Function PPTK_RowsFor(ByVal row As Long, ByVal col As Long) As Long
    Dim needed As Long
    If col < 1 Then col = 1
    needed = (ShapeCount + col - 1) \ col
    If row < needed Then row = needed
    PPTK_RowsFor = row
End Function

Private Function PPTK_Clamp(ByVal value As Long, ByVal low As Long, ByVal high As Long) As Long
    If high < low Then high = low
    If value < low Then value = low
    If value > high Then value = high
    PPTK_Clamp = value
End Function

Private Sub PPTK_EnsureCells()
    Dim row As Long
    Dim col As Long
    Dim name As String
    Dim box As MSForms.Label
    For row = 1 To VisRows
        For col = 1 To VisCols
            If Slots(row, col) Is Nothing Then
                name = "PPTKCell_" & Format$(row, "00") & "_" & Format$(col, "00")
                Set box = Nothing
                On Error Resume Next
                Set box = Me.Controls(name)
                On Error GoTo 0
                If box Is Nothing Then Set box = Me.Controls.Add("Forms.Label.1", name, True)
                box.Caption = ""
                box.Width = CellSize
                box.Height = CellSize
                box.BorderStyle = fmBorderStyleSingle
                box.SpecialEffect = fmSpecialEffectFlat
                box.BackStyle = fmBackStyleOpaque
                Set Slots(row, col) = New PPTKMatrixCell
                Slots(row, col).Attach Me, box, row, col
            End If
        Next col
    Next row
End Sub

Private Sub PPTK_StyleChrome()
    PPTK_StyleForm Me
    Me.BackColor = ColorPage
    Me.Font.Name = "Segoe UI"
    lblAccent.Visible = False   ' no coloured line under the title (dark PPTK look)
    PPTK_Font lblGuide, 9, False, ColorMuted
    lblGuide.BackStyle = fmBackStyleTransparent
    lblGuide.Caption = "Move over the grid to preview. Click to apply."
    PPTK_Font lblSize, 10, True, ColorInk
    lblSize.BackStyle = fmBackStyleTransparent
    lblSize.TextAlign = fmTextAlignCenter
    PPTK_Font lblError, 8, False, ColorMuted
    lblError.BackStyle = fmBackStyleTransparent
    lblError.TextAlign = fmTextAlignCenter
    PPTK_Button cmdCancel, False
    PPTK_Button cmdMore, False
    cmdCancel.Caption = "Cancel"
    cmdMore.Visible = False
    Set flatCancel = New PPTKFlatButton
    flatCancel.Attach Me, cmdCancel, False
End Sub

Private Sub PPTK_Layout()
    Dim row As Long
    Dim col As Long
    Dim gridW As Single
    Dim gridH As Single
    Dim inner As Single
    Dim origin As Single
    Dim topGrid As Single
    gridW = VisCols * CellSize + (VisCols - 1) * CellGap
    gridH = VisRows * CellSize + (VisRows - 1) * CellGap
    inner = gridW + (Pad * 2)
    If inner < 268 Then inner = 268
    origin = (inner - gridW) / 2
    topGrid = 32
    lblGuide.Left = Pad
    lblGuide.Top = 10
    lblGuide.Width = inner - (Pad * 2)
    lblGuide.Height = 16
    For row = 1 To 30
        For col = 1 To 30
            If Not Slots(row, col) Is Nothing Then
                With Slots(row, col).Box
                    .Visible = (row <= VisRows And col <= VisCols)
                    .Left = origin + (col - 1) * (CellSize + CellGap)
                    .Top = topGrid + (row - 1) * (CellSize + CellGap)
                    .Width = CellSize
                    .Height = CellSize
                End With
            End If
        Next col
    Next row
    lblSize.Left = Pad
    lblSize.Top = topGrid + gridH + 10
    lblSize.Width = inner - (Pad * 2)
    lblSize.Height = 18
    lblError.Left = Pad
    lblError.Top = lblSize.Top + 18
    lblError.Width = lblSize.Width
    lblError.Height = 14
    cmdCancel.Top = lblError.Top + 20
    cmdCancel.Height = 26
    cmdCancel.Width = 76
    cmdCancel.Left = inner - Pad - cmdCancel.Width
    Me.Width = inner + 16
    Me.Height = cmdCancel.Top + cmdCancel.Height + 44
    If Not flatCancel Is Nothing Then flatCancel.Sync
End Sub

Private Sub PPTK_Paint()
    Dim row As Long
    Dim col As Long
    Dim rows As Long
    Dim index As Long
    rows = PPTK_RowsFor(HoverRow, HoverCol)
    For row = 1 To VisRows
        For col = 1 To VisCols
            index = (row - 1) * HoverCol + col
            If row = HoverRow And col = HoverCol Then
                If index <= ShapeCount Then
                    Slots(row, col).Paint ColorFilled, ColorHoverLine
                Else
                    Slots(row, col).Paint ColorHover, ColorHoverLine
                End If
            ElseIf col <= HoverCol And row <= rows And index <= ShapeCount Then
                Slots(row, col).Paint ColorFilled, ColorFilledLine
            ElseIf col <= HoverCol And row <= rows Then
                Slots(row, col).Paint ColorCard, ColorLine
            Else
                Slots(row, col).Paint ColorDim, ColorDimLine
            End If
        Next col
    Next row
    lblSize.Caption = HoverCol & " columns " & ChrW$(215) & " " & rows & " rows"
    If rows * HoverCol > ShapeCount Then
        lblError.Caption = ShapeCount & " shapes, " & (rows * HoverCol - ShapeCount) & " empty cells"
    Else
        lblError.Caption = ShapeCount & " shapes"
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
