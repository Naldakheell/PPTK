VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTKSlideLibraryForm
   Caption         =   "Slide Library"
   ClientHeight    =   6400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840
   StartUpPosition =   1  'CenterOwner
   BackColor       =   &00FBF9F7&
   Begin MSForms.Label lblAccent
      BackColor       =   &00ECBA23&
      Height          =   80
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9840
   End
   Begin MSForms.TextBox txtSearch
      Height          =   420
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4560
   End
   Begin MSForms.ListBox lstEntries
      Height          =   3900
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   4560
   End
   Begin MSForms.Label lblEmpty
      Caption         =   "No matching entries"
      Height          =   360
      Left            =   360
      TabIndex        =   3
      Top             =   2200
      Visible         =   0   'False
      Width           =   4200
   End
   Begin MSForms.Image imgPreview
      BackColor       =   &00FFFFFF&
      BorderColor     =   &00EDE8E3&
      BorderStyle     =   1
      Height          =   2460
      Left            =   5040
      PictureSizeMode =   3
      TabIndex        =   4
      Top             =   780
      Width           =   4560
   End
   Begin MSForms.Label lblPreview
      Caption         =   "No preview"
      Height          =   360
      Left            =   5160
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   4200
   End
   Begin MSForms.Label lblName
      Caption         =   "Entry"
      Height          =   480
      Left            =   5040
      TabIndex        =   6
      Top             =   3360
      Width           =   4560
   End
   Begin MSForms.CheckBox chkSource
      Caption         =   "Insert with source formatting"
      Height          =   360
      Left            =   5040
      TabIndex        =   7
      Top             =   3900
      Width           =   4560
   End
   Begin MSForms.Label lblFolder
      Caption         =   "Folder"
      Height          =   480
      Left            =   240
      TabIndex        =   8
      Top             =   4860
      Width           =   9360
   End
   Begin MSForms.CommandButton btnInsert
      Caption         =   "Insert"
      Height          =   420
      Left            =   5040
      TabIndex        =   9
      Top             =   5640
      Width           =   1440
   End
   Begin MSForms.CommandButton btnFolder
      Caption         =   "Change Folder"
      Height          =   420
      Left            =   6600
      TabIndex        =   10
      Top             =   5640
      Width           =   1800
   End
   Begin MSForms.CommandButton btnClose
      Caption         =   "Close"
      Height          =   420
      Left            =   8520
      TabIndex        =   11
      Top             =   5640
      Width           =   1080
   End
End
Attribute VB_Name = "PPTKSlideLibraryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EntryNames() As String
Private EntryPaths() As String

Private Sub UserForm_Initialize()
    PPTK_StyleForm Me
    Me.BackColor = PPTK_COLOR_LIGHT_BACKGROUND
    Me.Font.Name = "Segoe UI"
    lblAccent.Caption = ""
    lblAccent.BackColor = PPTK_COLOR_PRIMARY
    PPTK_StyleField txtSearch
    PPTK_StyleField lstEntries
    PPTK_StyleField imgPreview
    PPTK_StylePrimaryButton btnInsert
    PPTK_StyleSecondaryButton btnFolder
    PPTK_StyleSecondaryButton btnClose
    lblName.BackStyle = 0
    lblName.ForeColor = PPTK_COLOR_DARK_TEXT
    lblFolder.BackStyle = 0
    lblFolder.ForeColor = PPTK_COLOR_DARK_TEXT
    lblEmpty.BackStyle = 0
    lblEmpty.ForeColor = PPTK_COLOR_DARK_TEXT
    lblPreview.BackStyle = 0
    lblPreview.ForeColor = PPTK_COLOR_DARK_TEXT
    chkSource.BackColor = PPTK_COLOR_LIGHT_BACKGROUND
    chkSource.ForeColor = PPTK_COLOR_DARK_TEXT
    lstEntries.ColumnCount = 2
    lstEntries.ColumnWidths = "180 pt;0 pt"
    txtSearch.Text = ""
    PPTK_Refill
End Sub

Private Sub txtSearch_Change()
    PPTK_Refill
End Sub

Private Sub lstEntries_Click()
    PPTK_ShowSelected
End Sub

Private Sub btnInsert_Click()
    PPTK_InsertSelected
End Sub

Private Sub lstEntries_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    PPTK_InsertSelected
End Sub

Private Sub btnFolder_Click()
    If PPTK_LibraryPickFolder() Then PPTK_Refill
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Unload Me
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    Set imgPreview.Picture = Nothing
End Sub

Public Sub PPTK_ApplySearch(ByVal query As String)
    txtSearch.Text = query
End Sub

Public Function PPTK_ListedCount() As Long
    PPTK_ListedCount = lstEntries.ListCount
End Function

Public Function PPTK_ListedName(ByVal index As Long) As String
    If index < 0 Or index >= lstEntries.ListCount Then Exit Function
    PPTK_ListedName = CStr(lstEntries.List(index, 0))
End Function

Private Sub PPTK_Refill()
    Dim query As String
    Dim index As Long
    Dim visibleCount As Long
    query = LCase$(Trim$(txtSearch.Text))
    PPTK_LibraryList EntryNames, EntryPaths
    lstEntries.Clear
    visibleCount = 0
    If (Not Not EntryNames) <> 0 Then
        If UBound(EntryNames) >= LBound(EntryNames) And UBound(EntryNames) > 0 Then
            For index = LBound(EntryNames) To UBound(EntryNames)
                If Len(query) = 0 Or InStr(1, LCase$(EntryNames(index)), query, vbTextCompare) > 0 Then
                    lstEntries.AddItem EntryNames(index)
                    lstEntries.List(lstEntries.ListCount - 1, 1) = EntryPaths(index)
                    visibleCount = visibleCount + 1
                End If
            Next index
        End If
    End If
    lblFolder.Caption = PPTK_LibraryFolder()
    lblEmpty.Visible = (visibleCount = 0)
    btnInsert.Enabled = (visibleCount > 0)
    If visibleCount > 0 Then
        lstEntries.ListIndex = 0
    Else
        lblName.Caption = "No entry selected"
        Set imgPreview.Picture = Nothing
        lblPreview.Visible = True
        lblPreview.Caption = "No preview"
    End If
End Sub

Private Sub PPTK_ShowSelected()
    Dim libraryPath As String
    Dim previewPath As String
    If lstEntries.ListIndex < 0 Then Exit Sub
    libraryPath = CStr(lstEntries.List(lstEntries.ListIndex, 1))
    lblName.Caption = CStr(lstEntries.List(lstEntries.ListIndex, 0))
    previewPath = ""
    If LCase$(Right$(libraryPath, 5)) = ".pptx" Then
        previewPath = Left$(libraryPath, Len(libraryPath) - 5) & ".jpg"
    End If
    On Error Resume Next
    Set imgPreview.Picture = Nothing
    Err.Clear
    If Len(previewPath) > 0 Then
        If Len(Dir$(previewPath)) > 0 Then
            Set imgPreview.Picture = LoadPicture(previewPath)
        End If
    End If
    lblPreview.Visible = (Err.Number <> 0 Or imgPreview.Picture Is Nothing)
    If lblPreview.Visible Then lblPreview.Caption = "No preview"
    Err.Clear
End Sub

Private Sub PPTK_InsertSelected()
    Dim libraryPath As String
    If lstEntries.ListIndex < 0 Then Exit Sub
    libraryPath = CStr(lstEntries.List(lstEntries.ListIndex, 1))
    PPTK_LibraryInsert libraryPath, (chkSource.Value = True)
End Sub
