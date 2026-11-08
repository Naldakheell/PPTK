Attribute VB_Name = "ModulePPTKBrand"
Option Explicit

' PPTK brand colors for UserForms.
' Long values are VBA RGB() colors: red + green * 256 + blue * 65536.
' Imported in the Windows project. The three PPTK forms call these helpers.
' Do not import this module into the Mac workbook.

Public Const PPTK_COLOR_PRIMARY As Long = 15514147
Public Const PPTK_COLOR_WHITE As Long = 16777215
Public Const PPTK_COLOR_DARK_TEXT As Long = 3879984
Public Const PPTK_COLOR_LIGHT_BACKGROUND As Long = 16513527
Public Const PPTK_COLOR_SUBTLE_BORDER As Long = 15591651

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

Public Function PPTK_BrandRgb(ByVal redValue As Long, ByVal greenValue As Long, ByVal blueValue As Long) As Long
    PPTK_BrandRgb = RGB(redValue, greenValue, blueValue)
End Function

Public Sub PPTK_StyleForm(ByVal targetForm As Object)
    If targetForm Is Nothing Then Exit Sub
    On Error Resume Next
    targetForm.BackColor = PPTK_COLOR_WHITE
    targetForm.ForeColor = PPTK_COLOR_DARK_TEXT
End Sub

Public Sub PPTK_StylePrimaryButton(ByVal button As Object)
    If button Is Nothing Then Exit Sub
    On Error Resume Next
    button.BackStyle = 1
    button.BackColor = PPTK_COLOR_PRIMARY
    button.ForeColor = PPTK_COLOR_WHITE
    button.BorderStyle = 0
    button.SpecialEffect = 0
End Sub

Public Sub PPTK_StyleSecondaryButton(ByVal button As Object)
    If button Is Nothing Then Exit Sub
    On Error Resume Next
    button.BackStyle = 1
    button.BackColor = PPTK_COLOR_LIGHT_BACKGROUND
    button.ForeColor = PPTK_COLOR_DARK_TEXT
    button.BorderStyle = 1
    button.BorderColor = PPTK_COLOR_SUBTLE_BORDER
    button.SpecialEffect = 0
End Sub

Public Sub PPTK_StyleField(ByVal field As Object)
    If field Is Nothing Then Exit Sub
    On Error Resume Next
    field.BackStyle = 1
    field.BackColor = PPTK_COLOR_WHITE
    field.ForeColor = PPTK_COLOR_DARK_TEXT
    field.BorderStyle = 1
    field.BorderColor = PPTK_COLOR_SUBTLE_BORDER
    field.SpecialEffect = 0
End Sub
