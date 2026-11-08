VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPTKMatrixCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One square in the Align Matrix picker. The form keeps every cell alive.

Private WithEvents Face As MSForms.Label
Private Host As Object
Private CellRow As Long
Private CellCol As Long

Public Sub Attach(ByVal owner As Object, ByVal box As MSForms.Label, ByVal row As Long, ByVal col As Long)
    Set Host = owner
    Set Face = box
    CellRow = row
    CellCol = col
End Sub

Public Function Box() As MSForms.Label
    Set Box = Face
End Function

Public Sub Paint(ByVal fill As Long, ByVal edge As Long)
    Face.BackColor = fill
    Face.BorderColor = edge
End Sub

Private Sub Face_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Host.PPTK_Hover CellRow, CellCol
End Sub

Private Sub Face_Click()
    Host.PPTK_Choose CellRow, CellCol
End Sub
