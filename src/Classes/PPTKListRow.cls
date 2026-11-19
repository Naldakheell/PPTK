VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPTKListRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One visible row of a PPTKDarkList. Passes clicks, double-clicks and mouse moves to the list.

Private WithEvents Face As MSForms.Label
Private Owner As Object
Private Slot As Long

Public Sub Attach(ByVal list As Object, ByVal row As MSForms.Label, ByVal index As Long)
    Set Owner = list
    Set Face = row
    Slot = index
End Sub

Public Function RowLabel() As MSForms.Label
    Set RowLabel = Face
End Function

Private Sub Face_Click()
    Owner.RowClick Slot
End Sub

Private Sub Face_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    Owner.RowDblClick Slot
End Sub

Private Sub Face_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Owner.RowHover Slot
End Sub
