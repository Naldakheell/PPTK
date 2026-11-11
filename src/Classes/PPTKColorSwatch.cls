VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PPTKColorSwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Click sink for a label in Find & Replace Colors: a found color ("found"), a
' New Color quick pick ("pick"), a scope chip ("scope"), a color-type chip
' ("type"), or a flat button ("button"). The form owns every instance.

Private WithEvents target As MSForms.Label
Attribute target.VB_VarHelpID = -1
Private ParentForm As Object
Private swatchKind As String
Private swatchIndex As Long

Public Sub Initialize(ByVal labelControl As MSForms.Label, ByVal owner As Object, ByVal kind As String, ByVal index As Long)
    Set target = labelControl
    Set ParentForm = owner
    swatchKind = kind
    swatchIndex = index
End Sub

Private Sub Target_Click()
    If ParentForm Is Nothing Then Exit Sub
    ParentForm.PPTK_SwatchClick swatchKind, swatchIndex
End Sub
