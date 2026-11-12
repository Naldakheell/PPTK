Attribute VB_Name = "ModulePPTKSmartBar"
Option Explicit

' Command palette metadata for the modeless Smart Bar.
' Every entry stores a PPTK_DispatchShortcut token. Commands with no
' keyboard shortcut use an internal CMD token that PPTKKeys never sends.
' The form filters this list. It does not contain a second implementation.
'
' No Win32 API declarations. The Defender ASR rule "Block Win32 API calls
' from Office macros" (92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B) is enforced on
' managed PCs and refuses to save or load a PPTM/PPAM whose VBA declares
' user32 functions. Hide then Show brings the modeless palette to the front.

Public Sub PPTK_ShowSmartBar()
    On Error GoTo Clean
    ' Always a fresh palette. A hidden palette shown again keeps its old focus
    ' state: arrows work, but typed characters do not reach the search field
    ' until it is clicked. A new form starts with the caret in the search field.
    PPTK_UnloadSmartBar
    DoEvents
    PPTKSmartBarForm.PPTK_Open
    PPTKSmartBarForm.Show vbModeless
    On Error Resume Next
    PPTKSmartBarForm.txtFilter.SetFocus
Clean:
End Sub

Private Sub PPTK_UnloadSmartBar()
    Dim loadedForm As Object
    For Each loadedForm In VBA.UserForms
        If TypeName(loadedForm) = "PPTKSmartBarForm" Then
            Unload loadedForm
            Exit For
        End If
    Next loadedForm
End Sub

Public Sub PPTK_HideSmartBar()
    On Error Resume Next
    If PPTKSmartBarForm.Visible Then PPTKSmartBarForm.Hide
End Sub

Public Function PPTK_CommandCount() As Long
    PPTK_CommandCount = 60
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
        Case 6: PPTK_Set commandName, shortcut, token, aliases, "Add Right Column - Keep Bounds", "Ctrl+2", "CTRL 2", "column add insert keep bounds footprint default grid table"
        Case 7: PPTK_Set commandName, shortcut, token, aliases, "Add Bottom Row - Keep Bounds", "Ctrl+3", "CTRL 3", "row add insert keep bounds footprint default grid table"
        Case 8: PPTK_Set commandName, shortcut, token, aliases, "Delete Right Column - Keep Bounds", "Alt+2", "ALT 2", "column delete remove keep bounds footprint default grid table"
        Case 9: PPTK_Set commandName, shortcut, token, aliases, "Delete Bottom Row - Keep Bounds", "Alt+3", "ALT 3", "row delete remove keep bounds footprint default grid table"
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
        Case 31: PPTK_Set commandName, shortcut, token, aliases, "Match Corner Roundness", "A A", "A A", "corner radius rounded rectangle roundness match"
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
        Case 48: PPTK_Set commandName, shortcut, token, aliases, "Remove Margins", "Alt+.", "ALT .", "margin padding inset zero text table cell fit shrink tight"
        Case 49: PPTK_Set commandName, shortcut, token, aliases, "Increase Font Size", "Ctrl+]", "CTRL ]", "font"
        Case 50: PPTK_Set commandName, shortcut, token, aliases, "Decrease Font Size", "Ctrl+[", "CTRL [", "font"
        Case 51: PPTK_Set commandName, shortcut, token, aliases, "Bring Forward", "Ctrl+Shift+]", "CTRL SHIFT ]", "order"
        Case 52: PPTK_Set commandName, shortcut, token, aliases, "Send Backward", "Ctrl+Shift+[", "CTRL SHIFT [", "order"
        Case 53: PPTK_Set commandName, shortcut, token, aliases, "Add Right Column - Expand Bounds", "", "CMD ADD COLUMN EXPAND", "column add insert expand grow outward legacy grid table"
        Case 54: PPTK_Set commandName, shortcut, token, aliases, "Add Bottom Row - Expand Bounds", "", "CMD ADD ROW EXPAND", "row add insert expand grow outward legacy grid table"
        Case 55: PPTK_Set commandName, shortcut, token, aliases, "Delete Right Column - Shrink Bounds", "", "CMD DELETE COLUMN SHRINK", "column delete remove shrink legacy grid table"
        Case 56: PPTK_Set commandName, shortcut, token, aliases, "Delete Bottom Row - Shrink Bounds", "", "CMD DELETE ROW SHRINK", "row delete remove shrink legacy grid table"
        Case 57: PPTK_Set commandName, shortcut, token, aliases, "Flip Current Slide", "", "CMD FLIP SLIDE", "mirror slide horizontal flip rtl ltr layout"
        Case 58: PPTK_Set commandName, shortcut, token, aliases, "Flip Entire Presentation", "", "CMD FLIP PRESENTATION", "mirror presentation deck all slides horizontal flip rtl ltr layout"
        Case 59: PPTK_Set commandName, shortcut, token, aliases, "Unify Shape Type", "X T", "X T", "shape type autoshape change same match"
        Case 60: PPTK_Set commandName, shortcut, token, aliases, "Find & Replace Colors", "", "CMD REPLACE COLORS", "color colour colors replace find swap recolor hex palette fill outline font edit"
    End Select
End Sub

Private Sub PPTK_Set(ByRef commandName As String, ByRef shortcut As String, ByRef token As String, ByRef aliases As String, ByVal nameValue As String, ByVal shortcutValue As String, ByVal tokenValue As String, ByVal aliasValue As String)
    commandName = nameValue
    shortcut = shortcutValue
    token = tokenValue
    aliases = aliasValue
End Sub
