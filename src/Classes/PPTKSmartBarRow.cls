VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPTKSmartBarRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Click and hover sink for one Smart Bar result label or icon image.
' The form owns every instance so the label can be removed without a stale sink.

Private WithEvents target As MSForms.Label
Attribute target.VB_VarHelpID = -1
Private WithEvents targetImage As MSForms.Image
Attribute targetImage.VB_VarHelpID = -1
Private ParentForm As Object
Private rowIndex As Long

Public Sub Initialize(ByVal labelControl As MSForms.Label, ByVal owner As Object, ByVal index As Long)
    Set target = labelControl
    Set ParentForm = owner
    rowIndex = index
End Sub

Public Sub InitializeImage(ByVal imageControl As MSForms.Image, ByVal owner As Object, ByVal index As Long)
    Set targetImage = imageControl
    Set ParentForm = owner
    rowIndex = index
End Sub

Private Sub TargetImage_Click()
    If ParentForm Is Nothing Then Exit Sub
    ParentForm.PPTK_RowClick rowIndex
End Sub

Private Sub TargetImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If ParentForm Is Nothing Then Exit Sub
    ParentForm.PPTK_RowHover rowIndex
End Sub

Private Sub Target_Click()
    If ParentForm Is Nothing Then Exit Sub
    ParentForm.PPTK_RowClick rowIndex
End Sub

Private Sub Target_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If ParentForm Is Nothing Then Exit Sub
    ParentForm.PPTK_RowHover rowIndex
End Sub

Private Sub Class_Terminate()
    Set target = Nothing
    Set targetImage = Nothing
    Set ParentForm = Nothing
End Sub
