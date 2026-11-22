Attribute VB_Name = "ModulePPTKRibbon"
Option Explicit

' One ribbon entry for every PPTK shortcut button.
' The button Tag is the token passed to PPTK_DispatchShortcut.

Public Sub PPTK_RibbonShortcut(ByVal control As IRibbonControl)
    If control Is Nothing Then Exit Sub
    PPTK_DispatchShortcut control.Tag
End Sub

' Direct ribbon callbacks for commands that have no keyboard token.

Public Sub PPTK_RibbonAddRightColumnExpand(ByVal control As IRibbonControl)
    PPTK_AddRightColumnExpand
End Sub

Public Sub PPTK_RibbonAddBottomRowExpand(ByVal control As IRibbonControl)
    PPTK_AddBottomRowExpand
End Sub

Public Sub PPTK_RibbonDeleteRightColumnShrink(ByVal control As IRibbonControl)
    PPTK_DeleteRightColumnShrink
End Sub

Public Sub PPTK_RibbonDeleteBottomRowShrink(ByVal control As IRibbonControl)
    PPTK_DeleteBottomRowShrink
End Sub

Public Sub PPTK_RibbonFlipCurrentSlide(ByVal control As IRibbonControl)
    PPTK_FlipCurrentSlide
End Sub

Public Sub PPTK_RibbonFlipEntirePresentation(ByVal control As IRibbonControl)
    PPTK_FlipEntirePresentation
End Sub
