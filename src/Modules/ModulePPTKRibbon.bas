Attribute VB_Name = "ModulePPTKRibbon"
Option Explicit

' One ribbon entry for every PPTK shortcut button.
' The button Tag is the token passed to PPTK_DispatchShortcut.

Public Sub PPTK_RibbonShortcut(ByVal control As IRibbonControl)
    If control Is Nothing Then Exit Sub
    PPTK_DispatchShortcut control.Tag
End Sub
