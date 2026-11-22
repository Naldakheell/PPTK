Attribute VB_Name = "ModulePPTKBrand"
Option Explicit

' PPTK brand colors for UserForms.
' Long values are VBA RGB() colors: red + green * 256 + blue * 65536.
' Imported in the Windows project. PPTKMatrixForm, PPTKCentralForm, PPTKSlideLibraryForm and PPTKFontsForm call these helpers.
' Do not import this module into the Mac workbook.

' PPTK release version (MAJOR.MINOR.PATCH). One number for the whole release:
' keep it equal to <Version> in src/PPTKKeys/windows/PPTKKeys.csproj.
Public Const PPTK_VERSION As String = "1.2.0"

Public Const PPTK_COLOR_PRIMARY As Long = 15514147
Public Const PPTK_COLOR_WHITE As Long = 16777215
Public Const PPTK_COLOR_DARK_TEXT As Long = 3879984
Public Const PPTK_COLOR_LIGHT_BACKGROUND As Long = 16513527
Public Const PPTK_COLOR_SUBTLE_BORDER As Long = 15591651

' Dark window palette, the Smart Bar's (PPTKSmartBarForm) and Find & Replace Colors' look, shared by every PPTK
' window. RGB in the comments.
Public Const PPTK_DARK_PAGE As Long = 2103313        ' #111820 window background
Public Const PPTK_DARK_SURFACE As Long = 3549980     ' #1C2B36 fields, secondary buttons, panels
Public Const PPTK_DARK_SELECTED As Long = 5522213    ' #254354 selected / highlighted surface
Public Const PPTK_DARK_INK As Long = 16381424        ' #F0F5F9 text
Public Const PPTK_DARK_MUTED As Long = 13484463      ' #AFC1CD secondary text
Public Const PPTK_DARK_LINE As Long = 5457460        ' #344653 borders and rules
Public Const PPTK_DARK_STRONG_LINE As Long = 6706494 ' #3E5566 stronger border
Public Const PPTK_DARK_ERROR As Long = 7039999       ' #FF6B6B error text
Public Const PPTK_DARK_DIM As Long = 8417372         ' #5C7080 disabled text
Public Const PPTK_DARK_HOVER As Long = 4207652       ' #243440 row under the mouse
Public Const PPTK_PRIMARY_HOT As Long = 15978331     ' #5BCFF3 primary button under the mouse

Public Const PPTK_R_PRIMARY As Long = 35
Public Const PPTK_G_PRIMARY As Long = 186
Public Const PPTK_B_PRIMARY As Long = 236
Public Const PPTK_R_WHITE As Long = 255
Public Const PPTK_G_WHITE As Long = 255
Public Const PPTK_B_WHITE As Long = 255
Public Const PPTK_R_DARK_TEXT As Long = 48
Public Const PPTK_G_DARK_TEXT As Long = 52
Public Const PPTK_B_DARK_TEXT As Long = 59
Public Const PPTK_R_LIGHT_BACKGROUND As Long = 247
Public Const PPTK_G_LIGHT_BACKGROUND As Long = 249
Public Const PPTK_B_LIGHT_BACKGROUND As Long = 251
Public Const PPTK_R_SUBTLE_BORDER As Long = 227
Public Const PPTK_G_SUBTLE_BORDER As Long = 232
Public Const PPTK_B_SUBTLE_BORDER As Long = 237

' The running add-in's version, for Application.Run callers (PPTKKeys, install
' checks). Read-only; no UI.
Public Function PPTK_GetVersion() As String
    PPTK_GetVersion = PPTK_VERSION
End Function

Public Function PPTK_BrandRgb(ByVal redValue As Long, ByVal greenValue As Long, ByVal blueValue As Long) As Long
    PPTK_BrandRgb = RGB(redValue, greenValue, blueValue)
End Function

' The four helpers below give every PPTK window the dark Smart Bar look (2026-11-18; before, light).
Public Sub PPTK_StyleForm(ByVal targetForm As Object)
    If targetForm Is Nothing Then Exit Sub
    On Error Resume Next
    targetForm.BackColor = PPTK_DARK_PAGE
    targetForm.ForeColor = PPTK_DARK_INK
End Sub

Public Sub PPTK_StylePrimaryButton(ByVal button As Object)
    If button Is Nothing Then Exit Sub
    On Error Resume Next
    button.BackStyle = 1
    button.BackColor = PPTK_COLOR_PRIMARY
    button.ForeColor = PPTK_DARK_PAGE
    button.BorderStyle = 0
    button.SpecialEffect = 0
End Sub

Public Sub PPTK_StyleSecondaryButton(ByVal button As Object)
    If button Is Nothing Then Exit Sub
    On Error Resume Next
    button.BackStyle = 1
    button.BackColor = PPTK_DARK_SURFACE
    button.ForeColor = PPTK_DARK_INK
    button.BorderStyle = 1
    button.BorderColor = PPTK_DARK_LINE
    button.SpecialEffect = 0
End Sub

Public Sub PPTK_StyleField(ByVal field As Object)
    If field Is Nothing Then Exit Sub
    On Error Resume Next
    field.BackStyle = 1
    field.BackColor = PPTK_DARK_SURFACE
    field.ForeColor = PPTK_DARK_INK
    field.BorderStyle = 1
    field.BorderColor = PPTK_DARK_LINE
    field.SpecialEffect = 0
End Sub
