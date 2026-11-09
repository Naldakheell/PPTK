Attribute VB_Name = "ModulePPTKSmartBar"
Option Explicit

' Command palette metadata for the modeless Smart Bar.
' Keyboard commands store a PPTK_DispatchShortcut token.
' Commands with no shortcut store PROC and a public procedure name.
' The form filters this list. It does not contain a second implementation.
' The palette is modeless, so PowerPoint would otherwise cover it.

Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal className As String, ByVal windowName As String) As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal insertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long

Public Sub PPTK_ShowSmartBar()
    On Error GoTo Clean
    If Not PPTKSmartBarForm.Visible Then
        PPTKSmartBarForm.txtFilter.Text = ""
        PPTKSmartBarForm.Show vbModeless
    End If
    On Error Resume Next
    PPTKSmartBarForm.txtFilter.SetFocus
    PPTK_RaiseSmartBar
Clean:
End Sub

Private Sub PPTK_RaiseSmartBar()
    Dim hwnd As LongPtr
    On Error Resume Next
    hwnd = FindWindowA("ThunderDFrame", PPTKSmartBarForm.Caption)
    If hwnd = 0 Then hwnd = FindWindowA(vbNullString, PPTKSmartBarForm.Caption)
    If hwnd = 0 Then Exit Sub
    SetWindowPos hwnd, -1, 0, 0, 0, 0, &H43
    SetForegroundWindow hwnd
End Sub

Public Sub PPTK_HideSmartBar()
    On Error Resume Next
    If PPTKSmartBarForm.Visible Then PPTKSmartBarForm.Hide
End Sub

Public Function PPTK_CommandCount() As Long
    PPTK_CommandCount = 59
End Function

Public Sub PPTK_CommandAt(ByVal index As Long, ByRef commandName As String, ByRef shortcut As String, ByRef token As String, ByRef aliases As String)
    commandName = ""
    shortcut = ""
    token = ""
    aliases = ""
    Select Case index
        Case 1: PPTK_Set commandName, shortcut, token, aliases, "Smart Bar", "Ctrl+K", "CTRL K", "search commands palette"
        Case 2: PPTK_Set commandName, shortcut, token, aliases, "Fix Slide Colors", "Q", "Q", "colors theme harmonize palette"
        Case 3: PPTK_Set commandName, shortcut, token, aliases, "Align Process", "A P", "A P", "process chevron pentagon"
        Case 4: PPTK_Set commandName, shortcut, token, aliases, "Align Matrix", "A X", "A X", "grid matrix rows columns"
        Case 5: PPTK_Set commandName, shortcut, token, aliases, "Align Central", "A Z", "A Z", "circle radial central"
        Case 6: PPTK_Set commandName, shortcut, token, aliases, "Add Column - Keep Bounds", "Ctrl+2", "CTRL 2", "column add bounds footprint"
        Case 7: PPTK_Set commandName, shortcut, token, aliases, "Add Row - Keep Bounds", "Ctrl+3", "CTRL 3", "row add bounds footprint"
        Case 8: PPTK_Set commandName, shortcut, token, aliases, "Delete Column - Keep Bounds", "Alt+2", "ALT 2", "column delete bounds footprint"
        Case 9: PPTK_Set commandName, shortcut, token, aliases, "Delete Row - Keep Bounds", "Alt+3", "ALT 3", "row delete bounds footprint"
        Case 10: PPTK_Set commandName, shortcut, token, aliases, "Add Right TK", "Ctrl+4", "CTRL 4", "takeaway tk right"
        Case 11: PPTK_Set commandName, shortcut, token, aliases, "Add Bottom TK", "Ctrl+5", "CTRL 5", "takeaway tk bottom"
        Case 12: PPTK_Set commandName, shortcut, token, aliases, "Add Top TK", "Ctrl+6", "CTRL 6", "takeaway tk top"
        Case 13: PPTK_Set commandName, shortcut, token, aliases, "Add Left TK", "Ctrl+7", "CTRL 7", "takeaway tk left"
        Case 14: PPTK_Set commandName, shortcut, token, aliases, "Insert Textbox", "T", "T", "text"
        Case 15: PPTK_Set commandName, shortcut, token, aliases, "Insert Rectangle", "R", "R", "shape"
        Case 16: PPTK_Set commandName, shortcut, token, aliases, "Insert Rounded Rectangle", "U", "U", "shape"
        Case 17: PPTK_Set commandName, shortcut, token, aliases, "Insert Oval", "O", "O", "circle ellipse"
        Case 18: PPTK_Set commandName, shortcut, token, aliases, "Insert Chevron", "C", "C", "process"
        Case 19: PPTK_Set commandName, shortcut, token, aliases, "Insert Pentagon", "P", "P", "process"
        Case 20: PPTK_Set commandName, shortcut, token, aliases, "Insert Triangle", "G", "G", "shape"
        Case 21: PPTK_Set commandName, shortcut, token, aliases, "Insert Parallelogram", "M", "M", "shape"
        Case 22: PPTK_Set commandName, shortcut, token, aliases, "Insert New Slide", "N", "N", "slide"
        Case 23: PPTK_Set commandName, shortcut, token, aliases, "Insert Icon", "I", "I", "icon"
        Case 24: PPTK_Set commandName, shortcut, token, aliases, "Enclose Icons", "E", "E", "circle icon"
        Case 25: PPTK_Set commandName, shortcut, token, aliases, "Align Left", "A L", "A L", "alignment"
        Case 26: PPTK_Set commandName, shortcut, token, aliases, "Align Right", "A R", "A R", "alignment"
        Case 27: PPTK_Set commandName, shortcut, token, aliases, "Align Top", "A T", "A T", "alignment"
        Case 28: PPTK_Set commandName, shortcut, token, aliases, "Align Bottom", "A B", "A B", "alignment"
        Case 29: PPTK_Set commandName, shortcut, token, aliases, "Align Middle", "A M", "A M", "alignment"
        Case 30: PPTK_Set commandName, shortcut, token, aliases, "Align Center", "A C", "A C", "alignment"
        Case 31: PPTK_Set commandName, shortcut, token, aliases, "Match Corner Roundness", "A A", "A A", "corner radius rounded rectangle"
        Case 32: PPTK_Set commandName, shortcut, token, aliases, "Same Width", "S W", "S W", "size"
        Case 33: PPTK_Set commandName, shortcut, token, aliases, "Same Height", "S H", "S H", "size"
        Case 34: PPTK_Set commandName, shortcut, token, aliases, "Same Size", "S S", "S S", "size"
        Case 35: PPTK_Set commandName, shortcut, token, aliases, "Distribute Horizontally", "D H", "D H", "space"
        Case 36: PPTK_Set commandName, shortcut, token, aliases, "Distribute Vertically", "D V", "D V", "space"
        Case 37: PPTK_Set commandName, shortcut, token, aliases, "Swap Shapes", "/", "/", "position"
        Case 38: PPTK_Set commandName, shortcut, token, aliases, "Swap Text", "Ctrl+/", "CTRL /", "text"
        Case 39: PPTK_Set commandName, shortcut, token, aliases, "Merge Textboxes", ". T", ". T", "text"
        Case 40: PPTK_Set commandName, shortcut, token, aliases, "Save Selected Slides", "\", "\", "save export"
        Case 41: PPTK_Set commandName, shortcut, token, aliases, "Increase Horizontal Spacing", "H +", "H +", "gap"
        Case 42: PPTK_Set commandName, shortcut, token, aliases, "Decrease Horizontal Spacing", "H -", "H -", "gap"
        Case 43: PPTK_Set commandName, shortcut, token, aliases, "Increase Vertical Spacing", "V +", "V +", "gap"
        Case 44: PPTK_Set commandName, shortcut, token, aliases, "Decrease Vertical Spacing", "V -", "V -", "gap"
        Case 45: PPTK_Set commandName, shortcut, token, aliases, "Copy Formatting", "Ctrl+Shift+C", "CTRL SHIFT C", "format painter"
        Case 46: PPTK_Set commandName, shortcut, token, aliases, "Paste Formatting", "Ctrl+Shift+V", "CTRL SHIFT V", "format painter"
        Case 47: PPTK_Set commandName, shortcut, token, aliases, "Paste Content Only", "Ctrl+Shift+T", "CTRL SHIFT T", "text"
        Case 48: PPTK_Set commandName, shortcut, token, aliases, "Remove Margins", "Alt+.", "ALT .", "margin"
        Case 49: PPTK_Set commandName, shortcut, token, aliases, "Increase Font Size", "Ctrl+]", "CTRL ]", "font"
        Case 50: PPTK_Set commandName, shortcut, token, aliases, "Decrease Font Size", "Ctrl+[", "CTRL [", "font"
        Case 51: PPTK_Set commandName, shortcut, token, aliases, "Bring Forward", "Ctrl+Shift+]", "CTRL SHIFT ]", "order"
        Case 52: PPTK_Set commandName, shortcut, token, aliases, "Send Backward", "Ctrl+Shift+[", "CTRL SHIFT [", "order"
        Case 53: PPTK_Set commandName, shortcut, token, aliases, "Add Column - Expand", "", "PROC PPTK_AddRightColumnExpand", "column add expand grow legacy"
        Case 54: PPTK_Set commandName, shortcut, token, aliases, "Add Row - Expand", "", "PROC PPTK_AddBottomRowExpand", "row add expand grow legacy"
        Case 55: PPTK_Set commandName, shortcut, token, aliases, "Delete Column - Shrink", "", "PROC PPTK_DeleteRightColumnShrink", "column delete shrink legacy"
        Case 56: PPTK_Set commandName, shortcut, token, aliases, "Delete Row - Shrink", "", "PROC PPTK_DeleteBottomRowShrink", "row delete shrink legacy"
        Case 57: PPTK_Set commandName, shortcut, token, aliases, "Flip Current Slide", "", "PROC PPTK_FlipCurrentSlide", "mirror slide horizontal flip"
        Case 58: PPTK_Set commandName, shortcut, token, aliases, "Flip Entire Presentation", "", "PROC PPTK_FlipEntirePresentation", "mirror presentation deck flip"
        Case 59: PPTK_Set commandName, shortcut, token, aliases, "Unify Shape Type", "X T", "X T", "shape type autoshape"
    End Select
End Sub

Private Sub PPTK_Set(ByRef commandName As String, ByRef shortcut As String, ByRef token As String, ByRef aliases As String, ByVal nameValue As String, ByVal shortcutValue As String, ByVal tokenValue As String, ByVal aliasValue As String)
    commandName = nameValue
    shortcut = shortcutValue
    token = tokenValue
    aliases = aliasValue
End Sub
