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
        PPTK_NewUndoEntry  ' Mac: the early-bound call does not compile on PowerPoint for Mac (see ModulePPTKShortcuts)
        selected.Align PPTK_AlignCommand(edge), msoTrue
        Exit Sub
    End If

    PPTK_NewUndoEntry  ' Mac: the early-bound call does not compile on PowerPoint for Mac (see ModulePPTKShortcuts)
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

' ---------------------------------------------------------------- distribute on one line (D H, D V)

' axis "H" (D H): the selected shapes line up on one row (Align Middle, same reference as A M) and are spaced
' evenly from left to right. axis "V" (D V): one column (Align Center, as A C), spaced evenly top to bottom.
' Works from two shapes. Order along the axis is the current position (left, or top); shapes at the same
' position are ordered by the other axis, so a column becomes a row and a row becomes a column.
' The outer shapes keep the span when everything fits in it; when the shapes do not fit (overlapping or
' stacked), they are laid out from the first one with an 8 pt gap. One shape: as before (to the slide).
Public Sub PPTK_DistributeLine(ByVal axis As String)
    Const OverlapGap As Single = 8
    Dim win As Object
    Dim inGroup As Boolean
    Dim selected As ShapeRange
    Dim order() As Shape
    Dim count As Long, index As Long, inner As Long
    Dim held As Shape
    Dim horizontal As Boolean
    Dim startPos As Single, endPos As Single, total As Single, gap As Single, position As Single

    axis = UCase$(Trim$(axis))
    If axis <> "H" And axis <> "V" Then Exit Sub
    horizontal = (axis = "H")
    On Error GoTo Clean
    Set win = Application.ActiveWindow
    If win.Selection.Type <> ppSelectionShapes Then Exit Sub
    inGroup = win.Selection.HasChildShapeRange
    If inGroup Then Set selected = win.Selection.ChildShapeRange Else Set selected = win.Selection.ShapeRange
    count = selected.Count
    If count < 2 Then
        ' One shape: centered on the slide (unchanged). One shape inside a group: nothing to do.
        If inGroup Then Exit Sub
        PPTK_NewUndoEntry  ' Mac: the early-bound call does not compile on PowerPoint for Mac (see ModulePPTKShortcuts)
        If horizontal Then selected.Align msoAlignCenters, msoTrue Else selected.Align msoAlignMiddles, msoTrue
        Exit Sub
    End If

    PPTK_NewUndoEntry  ' Mac: the early-bound call does not compile on PowerPoint for Mac (see ModulePPTKShortcuts)

    ' 1. One line, with the Align reference setting (first selected shape, or the selection bounds).
    If PPTK_AlignReference() = PPTK_ALIGN_BOUNDS Then
        If horizontal Then selected.Align msoAlignMiddles, msoFalse Else selected.Align msoAlignCenters, msoFalse
    Else
        For index = 2 To count
            If horizontal Then
                PPTK_AlignToReference selected(index), selected(1), "M"
            Else
                PPTK_AlignToReference selected(index), selected(1), "C"
            End If
        Next index
    End If

    ' 2. Order along the axis (insertion sort; selections are small).
    ReDim order(1 To count)
    For index = 1 To count
        Set order(index) = selected(index)
    Next index
    For index = 2 To count
        Set held = order(index)
        inner = index - 1
        Do While inner >= 1
            If Not PPTK_LineBefore(held, order(inner), horizontal) Then Exit Do
            Set order(inner + 1) = order(inner)
            inner = inner - 1
        Loop
        Set order(inner + 1) = held
    Next index

    ' 3. Even gaps between the outer edges; a fixed gap when the shapes do not fit.
    startPos = PPTK_LineStart(order(1), horizontal)
    endPos = startPos
    total = 0
    For index = 1 To count
        total = total + PPTK_LineSize(order(index), horizontal)
        If PPTK_LineStart(order(index), horizontal) + PPTK_LineSize(order(index), horizontal) > endPos Then
            endPos = PPTK_LineStart(order(index), horizontal) + PPTK_LineSize(order(index), horizontal)
        End If
    Next index
    gap = (endPos - startPos - total) / (count - 1)
    If gap < 0 Then gap = OverlapGap

    position = startPos
    For index = 1 To count
        If horizontal Then SetRealLeft order(index), position Else SetRealTop order(index), position
        position = position + PPTK_LineSize(order(index), horizontal) + gap
    Next index
Clean:
End Sub

' True when shape a comes before shape b along the axis (ties: the other axis).
Private Function PPTK_LineBefore(ByVal a As Shape, ByVal b As Shape, ByVal horizontal As Boolean) As Boolean
    Const Tolerance As Single = 0.5
    Dim mainA As Single, mainB As Single
    mainA = PPTK_LineStart(a, horizontal)
    mainB = PPTK_LineStart(b, horizontal)
    If Abs(mainA - mainB) > Tolerance Then
        PPTK_LineBefore = (mainA < mainB)
    Else
        PPTK_LineBefore = (PPTK_LineStart(a, Not horizontal) < PPTK_LineStart(b, Not horizontal) - Tolerance)
    End If
End Function

Private Function PPTK_LineStart(ByVal target As Shape, ByVal horizontal As Boolean) As Single
    If horizontal Then PPTK_LineStart = GetRealLeft(target) Else PPTK_LineStart = GetRealTop(target)
End Function

Private Function PPTK_LineSize(ByVal target As Shape, ByVal horizontal As Boolean) As Single
    If horizontal Then PPTK_LineSize = GetRealWidth(target) Else PPTK_LineSize = GetRealHeight(target)
End Function
