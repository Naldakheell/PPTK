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

' Slide Library window, dark PPTK look. The entry list is a PPTKDarkList (Smart Bar scrollbar) in the place
' of the designer ListBox lstEntries, which stays hidden. Click selects and previews, double-click inserts,
' Up / Down in the search box move the selection, Esc closes. The buttons are PPTKFlatButton faces over the
' real buttons.

Private EntryNames() As String
Private EntryPaths() As String
Private WithEvents lstLibrary As PPTKDarkList
Private WithEvents flatInsert As PPTKFlatButton
Private WithEvents flatFolder As PPTKFlatButton
Private WithEvents flatClose As PPTKFlatButton

Private Sub UserForm_Initialize()
    ' Dark PPTK window palette (ModulePPTKBrand, the Smart Bar's look).
    PPTK_StyleForm Me
    Me.BackColor = PPTK_DARK_PAGE
    Me.Font.Name = "Segoe UI"
    lblAccent.Visible = False   ' no coloured line under the title
    PPTK_StyleField txtSearch
    PPTK_StyleField imgPreview
    PPTK_StylePrimaryButton btnInsert
    PPTK_StyleSecondaryButton btnFolder
    PPTK_StyleSecondaryButton btnClose
    lblName.BackStyle = 0
    lblName.ForeColor = PPTK_DARK_INK
    lblFolder.BackStyle = 0
    lblFolder.ForeColor = PPTK_DARK_MUTED
    lblEmpty.Visible = False
    lblPreview.BackStyle = 0
    lblPreview.ForeColor = PPTK_DARK_MUTED
    chkSource.BackColor = PPTK_DARK_PAGE
    chkSource.ForeColor = PPTK_DARK_INK
    lstEntries.Visible = False
    lstEntries.TabStop = False
    Set lstLibrary = New PPTKDarkList
    lstLibrary.Init Me, "lstLibrary", lstEntries.Left, lstEntries.Top, lstEntries.Width, lstEntries.Height, 18
    lstLibrary.EmptyMessage = "No matching entries"
    Set flatInsert = New PPTKFlatButton
    flatInsert.Attach Me, btnInsert, True
    Set flatFolder = New PPTKFlatButton
    flatFolder.Attach Me, btnFolder, False
    Set flatClose = New PPTKFlatButton
    flatClose.Attach Me, btnClose, False
    txtSearch.Text = ""
    PPTK_Refill
End Sub

Private Sub txtSearch_Change()
    PPTK_Refill
End Sub

Private Sub lstLibrary_Changed()
    PPTK_ShowSelected
End Sub

Private Sub lstLibrary_Activated()
    PPTK_InsertSelected
End Sub

Private Sub btnInsert_Click()
    PPTK_InsertSelected
End Sub

Private Sub btnFolder_Click()
    If PPTK_LibraryPickFolder() Then PPTK_Refill
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub flatInsert_Click()
    btnInsert_Click
End Sub

Private Sub flatFolder_Click()
    btnFolder_Click
End Sub

Private Sub flatClose_Click()
    btnClose_Click
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    flatInsert.Leave
    flatFolder.Leave
    flatClose.Leave
    lstLibrary.Leave
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Unload Me
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
        If lstLibrary.HandleKey(KeyCode) Then KeyCode = 0
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
    PPTK_ListedCount = lstLibrary.Count
End Function

' index is 0-based (as the ListBox it replaces).
Public Function PPTK_ListedName(ByVal index As Long) As String
    If index < 0 Or index >= lstLibrary.Count Then Exit Function
    PPTK_ListedName = lstLibrary.ItemText(index + 1)
End Function

Private Sub PPTK_Refill()
    Dim query As String
    Dim index As Long
    Dim names() As String
    Dim paths() As String
    Dim visibleCount As Long
    query = LCase$(Trim$(txtSearch.Text))
    PPTK_LibraryList EntryNames, EntryPaths
    visibleCount = 0
    If (Not Not EntryNames) <> 0 Then
        If UBound(EntryNames) >= LBound(EntryNames) And UBound(EntryNames) > 0 Then
            ReDim names(1 To UBound(EntryNames) - LBound(EntryNames) + 1)
            ReDim paths(1 To UBound(names))
            For index = LBound(EntryNames) To UBound(EntryNames)
                If Len(query) = 0 Or InStr(1, LCase$(EntryNames(index)), query, vbTextCompare) > 0 Then
                    visibleCount = visibleCount + 1
                    names(visibleCount) = EntryNames(index)
                    paths(visibleCount) = EntryPaths(index)
                End If
            Next index
        End If
    End If
    If visibleCount > 0 Then
        ReDim Preserve names(1 To visibleCount)
        ReDim Preserve paths(1 To visibleCount)
        lstLibrary.SetItems names, paths
    Else
        lstLibrary.SetItems Empty
    End If
    lblFolder.Caption = PPTK_LibraryFolder()
    btnInsert.Enabled = (visibleCount > 0)
    flatInsert.Sync
    If visibleCount > 0 Then
        lstLibrary.ListIndex = 1
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
    If lstLibrary.ListIndex < 1 Then Exit Sub
    libraryPath = lstLibrary.ItemTag(lstLibrary.ListIndex)
    lblName.Caption = lstLibrary.ItemText(lstLibrary.ListIndex)
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
    If lstLibrary.ListIndex < 1 Then Exit Sub
    libraryPath = lstLibrary.ItemTag(lstLibrary.ListIndex)
    PPTK_LibraryInsert libraryPath, (chkSource.Value = True)
End Sub