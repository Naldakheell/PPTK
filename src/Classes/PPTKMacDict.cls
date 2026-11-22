Attribute VB_Name = "PPTKMacDict"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' PowerPoint for Mac has no Scripting.Dictionary. This class covers the part PPTK uses:
' CompareMode (keys always compare as text, like CompareMode = 1), Exists, Add, Keys and Count.
' Keys keep their first spelling and their insertion order.

Public CompareMode As Long
Private mKeys As Collection

Private Sub Class_Initialize()
    Set mKeys = New Collection
End Sub

Public Function Exists(ByVal key As Variant) As Boolean
    Dim probe As Variant
    On Error Resume Next
    probe = mKeys(LCase$(CStr(key)))
    Exists = (Err.Number = 0)
    Err.Clear
End Function

Public Sub Add(ByVal key As Variant, ByVal item As Variant)
    ' Raises 457 for a duplicate key, as Scripting.Dictionary does.
    mKeys.Add CStr(key), LCase$(CStr(key))
End Sub

Public Function Keys() As Variant
    Dim result() As Variant
    Dim index As Long
    If mKeys.Count = 0 Then
        Keys = Array()
        Exit Function
    End If
    ReDim result(0 To mKeys.Count - 1)
    For index = 1 To mKeys.Count
        result(index - 1) = mKeys(index)
    Next index
    Keys = result
End Function

Public Property Get Count() As Long
    Count = mKeys.Count
End Property
