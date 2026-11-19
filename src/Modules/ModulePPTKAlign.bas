Attribute VB_Name = "ModulePPTKAlign"
Option Explicit

' PPTK alignment: Align Left, Center, Right, Top, Middle and Bottom (A L, A C, A R, A T, A M, A B, and the
' Align ribbon buttons).
'
' Reference (setting PPTK\Align\Reference, changed with the ribbon check box "Align to selection" or the Smart
' Bar command "Align Reference"):
'   First (default)  The first selected shape is the reference. It does not move; every other selected shape
'                    moves so the same edge (or center) lines up with it. Employee feedback asked for this.
'   Bounds           PowerPoint's own Align: the edge of the whole selection (the previous PPTK behavior).
'
' Selection order, checked in PowerPoint on 2026-11-17: Selection.ShapeRange lists shapes in the order they
' were clicked (click, then Ctrl+click or Shift+click), also inside a group (ChildShapeRange), and changing the
' stacking order afterwards does not change it. A lasso (drag) selection lists the shapes back to front, so the
' reference is then the shape at the back.
'
' One selected shape is aligned to the slide, as before. Edges use the visible (rotated) bounding box
' (GetRealLeft/GetRealTop/GetRealWidth/GetRealHeight in ModuleObjects). One command is one undo step.
' No Win32 or Windows-only objects: this module can be used on the Mac as well.

Private Const SettingApp As String = "PPTK"
Private Const SettingSection As String = "Align"
Private Const SettingKey As String = "Reference"
Public Const PPTK_ALIGN_FIRST As String = "First"
Public Const PPTK_ALIGN_BOUNDS As String = "Bounds"
Private Const ToggleControlId As String = "PPTKAlignToSelection"

Private PPTK_Ribbon As IRibbonUI

' ---------------------------------------------------------------- reference setting

Public Function PPTK_AlignReference() As String
    Dim value As String
    value = GetSetting(SettingApp, SettingSection, SettingKey, PPTK_ALIGN_FIRST)
    If value <> PPTK_ALIGN_BOUNDS Then value = PPTK_ALIGN_FIRST
    PPTK_AlignReference = value
End Function

Public Sub PPTK_SetAlignReference(ByVal mode As String)
    If mode <> PPTK_ALIGN_BOUNDS Then mode = PPTK_ALIGN_FIRST
    SaveSetting SettingApp, SettingSection, SettingKey, mode
    PPTK_RefreshAlignToggle
End Sub

' Smart Bar "Align Reference" (token CMD ALIGN REFERENCE): switches the reference and says which one is now used.
Public Sub PPTK_ToggleAlignReference()
    If PPTK_AlignReference() = PPTK_ALIGN_FIRST Then
        PPTK_SetAlignReference PPTK_ALIGN_BOUNDS
        MsgBox "Align now lines shapes up with the edge of the whole selection.", vbInformation, "PPTK"
    Else
        PPTK_SetAlignReference PPTK_ALIGN_FIRST
        MsgBox "Align now lines shapes up with the first selected shape.", vbInformation, "PPTK"
    End If
End Sub

' ---------------------------------------------------------------- ribbon (check box "Align to selection")

Public Sub PPTK_RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set PPTK_Ribbon = ribbon
End Sub

' True once PowerPoint has loaded the PPTK ribbon (onLoad ran). Used by src\PPTK\tests\Test-Ribbon.ps1.
Public Function PPTK_RibbonLoaded() As Boolean
    PPTK_RibbonLoaded = Not (PPTK_Ribbon Is Nothing)
End Function

Public Sub PPTK_AlignToSelectionGetPressed(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = (PPTK_AlignReference() = PPTK_ALIGN_BOUNDS)
End Sub

Public Sub PPTK_AlignToSelectionClick(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then PPTK_SetAlignReference PPTK_ALIGN_BOUNDS Else PPTK_SetAlignReference PPTK_ALIGN_FIRST
End Sub

Private Sub PPTK_RefreshAlignToggle()
    On Error Resume Next
    If Not PPTK_Ribbon Is Nothing Then PPTK_Ribbon.InvalidateControl ToggleControlId
    Err.Clear
End Sub

' ---------------------------------------------------------------- alignment

' edge: "L" left, "C" center, "R" right, "T" top, "M" middle, "B" bottom.
Public Sub PPTK_AlignEdge(ByVal edge As String)
    Dim win As Object
    Dim inGroup As Boolean
    Dim selected As ShapeRange
    Dim reference As Shape
    Dim index As Long

    edge = UCase$(Trim$(edge))
    If PPTK_AlignCommand(edge) < 0 Then Exit Sub
    On Error GoTo Clean
    Set win = Application.ActiveWindow
    If win.Selection.Type <> ppSelectionShapes Then Exit Sub
    inGroup = win.Selection.HasChildShapeRange
    If inGroup Then Set selected = win.Selection.ChildShapeRange Else Set selected = win.Selection.ShapeRange
    If selected.Count < 1 Then Exit Sub

    If selected.Count = 1 Then
        ' One shape: to the slide (unchanged). One shape inside a group: nothing to line up with.
        If inGroup Then Exit Sub
        Application.StartNewUndoEntry
        selected.Align PPTK_AlignCommand(edge), msoTrue
        Exit Sub
    End If

    Application.StartNewUndoEntry
    If PPTK_AlignReference() = PPTK_ALIGN_BOUNDS Then
        selected.Align PPTK_AlignCommand(edge), msoFalse
        Exit Sub
    End If

    Set reference = selected(1)
    For index = 2 To selected.Count
        PPTK_AlignToReference selected(index), reference, edge
    Next index
Clean:
End Sub

' Moves one shape so its visible edge (or center) matches the reference's. The reference does not move.
Private Sub PPTK_AlignToReference(ByVal target As Shape, ByVal reference As Shape, ByVal edge As String)
    Select Case edge
        Case "L"
            SetRealLeft target, GetRealLeft(reference)
        Case "R"
            SetRealLeft target, GetRealLeft(reference) + GetRealWidth(reference) - GetRealWidth(target)
        Case "C"
            SetRealLeft target, GetRealLeft(reference) + (GetRealWidth(reference) - GetRealWidth(target)) / 2
        Case "T"
            SetRealTop target, GetRealTop(reference)
        Case "B"
            SetRealTop target, GetRealTop(reference) + GetRealHeight(reference) - GetRealHeight(target)
        Case "M"
            SetRealTop target, GetRealTop(reference) + (GetRealHeight(reference) - GetRealHeight(target)) / 2
    End Select
End Sub

Private Function PPTK_AlignCommand(ByVal edge As String) As Long
    Select Case edge
        Case "L": PPTK_AlignCommand = msoAlignLefts
        Case "C": PPTK_AlignCommand = msoAlignCenters
        Case "R": PPTK_AlignCommand = msoAlignRights
        Case "T": PPTK_AlignCommand = msoAlignTops
        Case "M": PPTK_AlignCommand = msoAlignMiddles
        Case "B": PPTK_AlignCommand = msoAlignBottoms
        Case Else: PPTK_AlignCommand = -1
    End Select
End Function
