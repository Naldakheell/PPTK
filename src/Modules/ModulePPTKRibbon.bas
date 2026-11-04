Attribute VB_Name = "ModulePPTKRibbon"
Option Explicit

' One Ribbon callback for the nine PPTK inserts. The control Tag is the token.

Public Sub PPTK_RibbonShortcut(control As IRibbonControl)
    PPTK_DispatchShortcut control.Tag
End Sub
