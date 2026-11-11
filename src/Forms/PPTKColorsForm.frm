VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTKColorsForm 
   Caption         =   "PPTK - Find & Replace Colors"
   ClientHeight    =   10800
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7460
   OleObjectBlob   =   "PPTKColorsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PPTKColorsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' PPTK Find & Replace Colors dialog, dark PPTK style (Smart Bar palette).
' Every control is created in code (PPTK_Build); the designer holds an empty
' form. Chips and flat buttons are labels with PPTKColorSwatch click sinks.
' Scanning and replacing live in ModulePPTKReplaceColors.
'
' Nothing changes until Replace is clicked and the summary is confirmed.
' Selecting a swatch never selects shapes or moves to a slide.

Private Const Bg As Long = 2103313             ' #111820
Private Const Surface As Long = 3549980        ' #1C2B36
Private Const Picked As Long = 5522213         ' #254354
Private Const Ink As Long = 16381424           ' #F0F5F9
Private Const Muted As Long = 11772815         ' #8FA3B3
Private Const Dim1 As Long = 8417372           ' #5C7080
Private Const Accent As Long = 15514147        ' #23BAEC
Private Const Rule As Long = 4799530           ' #2A3C49
Private Const ErrorRed As Long = 7039999       ' #FF6B6B

Private Const SwatchSize As Single = 34
Private Const CellWidth As Single = 44
Private Const CellHeight As Single = 50
Private Const GalleryCols As Long = 7

Private WithEvents txtFrom As MSForms.TextBox
Attribute txtFrom.VB_VarHelpID = -1
Private WithEvents txtTo As MSForms.TextBox
Attribute txtTo.VB_VarHelpID = -1
Private WithEvents txtHex As MSForms.TextBox
Attribute txtHex.VB_VarHelpID = -1
Private WithEvents btnEsc As MSForms.CommandButton
Attribute btnEsc.VB_VarHelpID = -1

Private fraGallery As MSForms.Frame
Private lblFoundCount As MSForms.label
Private lblEmpty As MSForms.label
Private lblHex As MSForms.label
Private lblUses As MSForms.label
Private lblSlides As MSForms.label
Private lblSelection As MSForms.label
Private lblPreview As MSForms.label
Private lblHexState As MSForms.label
Private lblSkip As MSForms.label
Private lblRangeDash As MSForms.label
Private ScopeBox(1 To 3) As MSForms.label
Private ScopeText(1 To 3) As MSForms.label
Private TypeBox(1 To 3) As MSForms.label
Private TypeText(1 To 3) As MSForms.label
Private ApplyBox As MSForms.label
Private ApplyText As MSForms.label
Private CloseBox As MSForms.label
Private CloseText As MSForms.label

Private Sinks As Collection
Private FixedSinks As Long
Private Rings() As MSForms.label
Private FoundRgb() As Long
Private IsPicked() As Boolean
Private FoundCount As Long
Private ActiveIndex As Long
Private ScopeMode As Long                       ' 1 all, 2 current, 3 range
Private TypeOn(1 To 3) As Boolean               ' fill, outline, font
Private ScopeFrom As Long
Private ScopeTo As Long
Private ScopeOk As Boolean
Private ApplyReady As Boolean
Private Building As Boolean
Private LastSummary As String
Private KeepPicked As String

Private Sub UserForm_Initialize()
    Building = True
    ScopeMode = 1
    TypeOn(1) = True
    PPTK_Build
    PPTK_StyleChips
    Building = False
    PPTK_Rescan
    PPTK_UpdateTarget
End Sub

' ---------- layout ----------

Private Sub PPTK_Build()
    Dim i As Long
    Dim pick As MSForms.label
    Dim captions As Variant
    Dim lefts As Variant
    Dim widths As Variant

    Me.caption = "PPTK - Find & Replace Colors"
    Me.BackColor = Bg
    Me.width = 352
    Me.height = 500
    Set Sinks = New Collection

    PPTK_Text "lblTitle", "Find & Replace Colors", 14, 10, 220, 20, 12.5, True, Ink
    With PPTK_Text("lblScopeNote", "Slide content only", 200, 15, 132, 12, 8, False, Dim1)
        .TextAlign = fmTextAlignRight
    End With

    ' Search scope chips and the range fields
    PPTK_Text "lblScope", "SEARCH", 14, 38, 100, 11, 7.5, True, Muted
    captions = Array("All slides", "Current slide", "Range")
    lefts = Array(14, 86, 170)
    widths = Array(68, 80, 54)
    For i = 1 To 3
        PPTK_Chip ScopeBox(i), ScopeText(i), "scope", i, captions(i - 1), lefts(i - 1), 51, widths(i - 1)
    Next i
    Set txtFrom = PPTK_Box("txtFrom", 232, 51, 40)
    Set lblRangeDash = PPTK_Text("lblRangeDash", ChrW$(&H2013), 274, 54, 14, 14, 9, False, Muted)
    lblRangeDash.TextAlign = fmTextAlignCenter
    Set txtTo = PPTK_Box("txtTo", 290, 51, 40)
    txtFrom.text = "1"
    txtTo.text = CStr(PPTK_ColorSlideTotal())

    ' Color type chips
    PPTK_Text "lblTypes", "COLOR TYPES", 14, 80, 100, 11, 7.5, True, Muted
    captions = Array("Fill", "Outline", "Font")
    lefts = Array(14, 64, 130)
    widths = Array(46, 62, 46)
    For i = 1 To 3
        PPTK_Chip TypeBox(i), TypeText(i), "type", i, captions(i - 1), lefts(i - 1), 93, widths(i - 1)
    Next i
    Set lblFoundCount = PPTK_Text("lblFoundCount", "", 180, 97, 152, 12, 8, False, Muted)
    lblFoundCount.TextAlign = fmTextAlignRight

    ' Gallery
    Set fraGallery = Me.Controls.Add("Forms.Frame.1", "fraGallery", True)
    With fraGallery
        .left = 14: .Top = 120: .width = 318: .height = 150
        .caption = ""
        .BackColor = Surface
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsNone
    End With
    Set lblEmpty = fraGallery.Controls.Add("Forms.Label.1", "lblEmpty", True)
    With lblEmpty
        .left = 14: .Top = 56: .width = 290: .height = 34
        .BackStyle = fmBackStyleTransparent
        .Font.name = "Segoe UI": .Font.size = 9: .Font.Weight = 400
        .ForeColor = Muted
        .TextAlign = fmTextAlignCenter
        .WordWrap = True
    End With

    ' Selected color details
    Set lblHex = PPTK_Text("lblHex", "", 14, 278, 80, 18, 11, True, Ink)
    Set lblUses = PPTK_Text("lblUses", "", 96, 281, 236, 14, 8.5, False, Muted)
    Set lblSlides = PPTK_Text("lblSlides", "", 14, 297, 318, 14, 8.5, False, Ink)
    Set lblSelection = PPTK_Text("lblSelection", "", 14, 313, 318, 14, 8.5, True, Accent)

    With Me.Controls.Add("Forms.Label.1", "lblRule", True)
        .left = 14: .Top = 334: .width = 318: .height = 1
        .BackStyle = fmBackStyleOpaque: .BackColor = Rule
    End With

    ' Replacement
    PPTK_Text "lblNew", "REPLACE WITH", 14, 343, 120, 11, 7.5, True, Muted
    Set lblPreview = PPTK_Text("lblPreview", "", 14, 357, 30, 22, 9, False, Ink)
    lblPreview.BackStyle = fmBackStyleOpaque
    lblPreview.BorderStyle = fmBorderStyleSingle
    lblPreview.BorderColor = Rule
    Set txtHex = PPTK_Box("txtHex", 50, 357, 78)
    txtHex.height = 22
    txtHex.text = "#"
    Set lblHexState = PPTK_Text("lblHexState", "", 134, 362, 198, 14, 8, False, Muted)
    For i = 1 To PPTK_PaletteCount()
        Set pick = Me.Controls.Add("Forms.Label.1", "pick_" & Format(i, "00"), True)
        With pick
            .left = 14 + (i - 1) * 21.15: .Top = 386: .width = 17: .height = 17
            .BackStyle = fmBackStyleOpaque
            .BackColor = PPTK_PaletteColor(i)
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = Rule
            .ControlTipText = PPTK_ColorHex(PPTK_PaletteColor(i))
        End With
        PPTK_Sink pick, "pick", i
    Next i

    Set lblSkip = PPTK_Text("lblSkip", "", 14, 410, 318, 22, 7.5, False, Dim1)
    lblSkip.WordWrap = True

    ' Flat buttons; a hidden native button keeps Esc working.
    PPTK_Chip CloseBox, CloseText, "button", 2, "Close", 176, 438, 74
    PPTK_Chip ApplyBox, ApplyText, "button", 1, "Replace", 256, 438, 76
    CloseBox.height = 26: CloseText.Top = CloseBox.Top + 6
    ApplyBox.height = 26: ApplyText.Top = ApplyBox.Top + 6
    CloseBox.BackColor = Surface: CloseBox.BorderColor = Rule: CloseText.ForeColor = Ink
    ApplyText.Font.Weight = 700
    Set btnEsc = Me.Controls.Add("Forms.CommandButton.1", "btnEsc", True)
    With btnEsc
        .left = -200: .Top = 0: .width = 10: .height = 10
        .Cancel = True
        .TabStop = False
    End With
    FixedSinks = Sinks.count
End Sub

Private Function PPTK_Text(ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal size As Single, ByVal isBold As Boolean, ByVal color As Long) As MSForms.label
    Dim lbl As MSForms.label
    Set lbl = Me.Controls.Add("Forms.Label.1", name, True)
    With lbl
        .left = x: .Top = y: .width = w: .height = h
        .caption = caption
        .BackStyle = fmBackStyleTransparent
        ' MSForms: Font.Bold = False leaves weight 700, so the weight is set directly.
        .Font.name = "Segoe UI": .Font.size = size: .Font.Weight = IIf(isBold, 700, 400)
        .ForeColor = color
        .WordWrap = False
    End With
    Set PPTK_Text = lbl
End Function

' A chip or flat button: a bordered box label and a centered caption label, both clickable.
Private Sub PPTK_Chip(ByRef box As MSForms.label, ByRef text As MSForms.label, ByVal kind As String, ByVal index As Long, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal w As Single)
    Set box = Me.Controls.Add("Forms.Label.1", kind & "_box_" & index, True)
    With box
        .left = x: .Top = y: .width = w: .height = 20
        .BackStyle = fmBackStyleOpaque
        .BackColor = Surface
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = Rule
    End With
    Set text = Me.Controls.Add("Forms.Label.1", kind & "_text_" & index, True)
    With text
        .left = x + 1: .Top = y + 4: .width = w - 2: .height = 13
        .caption = caption
        .BackStyle = fmBackStyleTransparent
        .Font.name = "Segoe UI": .Font.size = 8.5: .Font.Weight = 400
        .ForeColor = Muted
        .TextAlign = fmTextAlignCenter
    End With
    PPTK_Sink box, kind, index
    PPTK_Sink text, kind, index
End Sub

Private Sub PPTK_Sink(ByVal lbl As MSForms.label, ByVal kind As String, ByVal index As Long)
    Dim sink As PPTKColorSwatch
    Set sink = New PPTKColorSwatch
    sink.Initialize lbl, Me, kind, index
    Sinks.Add sink
End Sub

Private Function PPTK_Box(ByVal name As String, ByVal x As Single, ByVal y As Single, ByVal w As Single) As MSForms.TextBox
    Dim box As MSForms.TextBox
    Set box = Me.Controls.Add("Forms.TextBox.1", name, True)
    With box
        .left = x: .Top = y: .width = w: .height = 20
        .Font.name = "Segoe UI": .Font.size = 9: .Font.Weight = 400
        .BackColor = Surface
        .ForeColor = Ink
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = Rule
        .SpecialEffect = fmSpecialEffectFlat
    End With
    Set PPTK_Box = box
End Function

Private Sub PPTK_StyleChips()
    Dim i As Long
    For i = 1 To 3
        PPTK_StyleChip ScopeBox(i), ScopeText(i), (ScopeMode = i)
        PPTK_StyleChip TypeBox(i), TypeText(i), TypeOn(i)
    Next i
    txtFrom.Visible = (ScopeMode = 3)
    txtTo.Visible = (ScopeMode = 3)
    lblRangeDash.Visible = (ScopeMode = 3)
End Sub

Private Sub PPTK_StyleChip(ByVal box As MSForms.label, ByVal text As MSForms.label, ByVal isOn As Boolean)
    If isOn Then
        box.BackColor = Picked
        box.BorderColor = Accent
        text.ForeColor = Ink
    Else
        box.BackColor = Surface
        box.BorderColor = Rule
        text.ForeColor = Muted
    End If
End Sub

' ---------- clicks (PPTKColorSwatch) ----------

Public Sub PPTK_SwatchClick(ByVal kind As String, ByVal index As Long)
    Dim i As Long
    Select Case kind
        Case "pick"
            txtHex.text = PPTK_ColorHex(PPTK_PaletteColor(index))
        Case "scope"
            ScopeMode = index
            PPTK_StyleChips
            If index = 3 Then
                On Error Resume Next
                txtFrom.SetFocus
                On Error GoTo 0
            End If
            PPTK_Rescan
        Case "type"
            TypeOn(index) = Not TypeOn(index)
            PPTK_StyleChips
            PPTK_Rescan
        Case "button"
            If index = 1 Then
                If ApplyReady Then PPTK_Apply
            Else
                Unload Me
            End If
        Case "found"
            If index < 1 Or index > FoundCount Then Exit Sub
            IsPicked(index) = Not IsPicked(index)
            Rings(index).Visible = IsPicked(index)
            If IsPicked(index) Then
                ActiveIndex = index
            Else
                ActiveIndex = 0
                For i = FoundCount To 1 Step -1
                    If IsPicked(i) Then ActiveIndex = i
                Next i
            End If
            PPTK_UpdateInfo
    End Select
End Sub

Private Sub btnEsc_Click()
    Unload Me
End Sub

' ---------- scope and scan ----------

Private Sub txtFrom_Change()
    If Building Then Exit Sub
    If ScopeMode = 3 Then PPTK_Rescan
End Sub

Private Sub txtTo_Change()
    If Building Then Exit Sub
    If ScopeMode = 3 Then PPTK_Rescan
End Sub

' Resolves the scope; sets ScopeOk and returns a message when it is not valid.
Private Function PPTK_ResolveScope() As String
    Dim total As Long
    Dim a As Long
    Dim b As Long
    ScopeOk = False
    total = PPTK_ColorSlideTotal()
    PPTK_ResolveScope = ""
    If ScopeMode = 2 Then
        a = PPTK_ColorCurrentSlide()
        If a < 1 Then
            PPTK_ResolveScope = "Select a slide in the presentation to use Current slide."
            Exit Function
        End If
        b = a
    ElseIf ScopeMode = 3 Then
        If Not PPTK_WholeNumber(txtFrom.text, a) Or Not PPTK_WholeNumber(txtTo.text, b) Then
            PPTK_ResolveScope = "Enter slide numbers from 1 to " & total & "."
            Exit Function
        End If
        If a < 1 Or b > total Or a > b Then
            PPTK_ResolveScope = "Enter a slide range from 1 to " & total & ", with From not greater than To."
            Exit Function
        End If
    Else
        a = 1
        b = total
    End If
    ScopeFrom = a
    ScopeTo = b
    ScopeOk = True
End Function

Private Function PPTK_WholeNumber(ByVal text As String, ByRef value As Long) As Boolean
    Dim i As Long
    text = Trim$(text)
    PPTK_WholeNumber = False
    If Len(text) = 0 Or Len(text) > 5 Then Exit Function
    For i = 1 To Len(text)
        If InStr(1, "0123456789", Mid$(text, i, 1)) = 0 Then Exit Function
    Next i
    value = CLng(text)
    PPTK_WholeNumber = True
End Function

Private Sub PPTK_Rescan()
    Dim message As String
    Dim keep As String
    Dim i As Long
    If FoundCount > 0 Then
        keep = ","
        For i = 1 To FoundCount
            If IsPicked(i) Then keep = keep & FoundRgb(i) & ","
        Next i
        KeepPicked = keep
    Else
        keep = KeepPicked
        If Len(keep) = 0 Then keep = ","
    End If

    message = PPTK_ResolveScope()
    If Not ScopeOk Then
        lblSkip.caption = ""
        PPTK_ShowColors 0, keep, message
        Exit Sub
    End If
    If Not TypeOn(1) And Not TypeOn(2) And Not TypeOn(3) Then
        lblSkip.caption = ""
        PPTK_ShowColors 0, keep, "Turn on at least one color type: Fill, Outline, or Font."
        Exit Sub
    End If
    Me.MousePointer = fmMousePointerHourGlass
    PPTK_ColorScan ScopeFrom, ScopeTo, TypeOn(1), TypeOn(2), TypeOn(3)
    Me.MousePointer = fmMousePointerDefault
    PPTK_ShowColors PPTK_ColorCount(), keep, "No solid colors of the chosen types in " & PPTK_ScopeText() & "."
    lblSkip.caption = PPTK_ColorSkipText()
End Sub

Private Function PPTK_ScopeText() As String
    If ScopeFrom = ScopeTo Then
        PPTK_ScopeText = "slide " & ScopeFrom
    Else
        PPTK_ScopeText = "slides " & ScopeFrom & " to " & ScopeTo
    End If
End Function

' ---------- gallery ----------

Private Sub PPTK_ShowColors(ByVal count As Long, ByVal keep As String, ByVal emptyMessage As String)
    Dim i As Long
    Dim x As Single
    Dim y As Single
    Dim rgbValue As Long
    Dim total As Long
    Dim fills As Long
    Dim lines As Long
    Dim fonts As Long
    Dim slideList As String
    Dim swatch As MSForms.label
    Dim countLabel As MSForms.label

    For i = fraGallery.Controls.count - 1 To 0 Step -1
        If left$(fraGallery.Controls(i).name, 4) = "dyn_" Then fraGallery.Controls.Remove fraGallery.Controls(i).name
    Next i
    For i = Sinks.count To FixedSinks + 1 Step -1
        Sinks.Remove i
    Next i

    FoundCount = count
    ActiveIndex = 0
    If count > 0 Then
        ReDim FoundRgb(1 To count)
        ReDim IsPicked(1 To count)
        ReDim Rings(1 To count)
    End If
    lblEmpty.Visible = (count = 0)
    lblEmpty.caption = emptyMessage
    If count = 0 Then
        lblFoundCount.caption = ""
    Else
        lblFoundCount.caption = count & PPTK_Plural(count, " color", " colors") & "  " & ChrW$(&HB7) & "  most used first"
    End If

    For i = 1 To count
        PPTK_ColorInfo i, rgbValue, total, fills, lines, fonts, slideList
        FoundRgb(i) = rgbValue
        IsPicked(i) = (InStr(1, keep, "," & rgbValue & ",") > 0)
        x = 10 + ((i - 1) Mod GalleryCols) * CellWidth
        y = 10 + ((i - 1) \ GalleryCols) * CellHeight

        Set Rings(i) = fraGallery.Controls.Add("Forms.Label.1", "dyn_ring_" & i, True)
        With Rings(i)
            .left = x - 3: .Top = y - 3: .width = SwatchSize + 6: .height = SwatchSize + 6
            .BackStyle = fmBackStyleOpaque: .BackColor = Accent
            .Visible = IsPicked(i)
        End With
        Set swatch = fraGallery.Controls.Add("Forms.Label.1", "dyn_swatch_" & i, True)
        With swatch
            .left = x: .Top = y: .width = SwatchSize: .height = SwatchSize
            .BackStyle = fmBackStyleOpaque
            .BackColor = rgbValue
            .BorderStyle = fmBorderStyleSingle
            ' A lighter edge keeps dark colors (navy, black) visible on the dark gallery.
            .BorderColor = Dim1
            .ControlTipText = PPTK_ColorHex(rgbValue) & "  " & total & PPTK_Plural(total, " use", " uses")
        End With
        Set countLabel = fraGallery.Controls.Add("Forms.Label.1", "dyn_count_" & i, True)
        With countLabel
            .left = x - 4: .Top = y + SwatchSize + 2: .width = SwatchSize + 8: .height = 11
            .caption = CStr(total)
            .BackStyle = fmBackStyleTransparent
            .Font.name = "Segoe UI": .Font.size = 7.5: .Font.Weight = 700
            .ForeColor = Muted
            .TextAlign = fmTextAlignCenter
        End With
        PPTK_Sink swatch, "found", i
        PPTK_Sink countLabel, "found", i
    Next i

    fraGallery.ScrollHeight = 10 + ((count + GalleryCols - 1) \ GalleryCols) * CellHeight
    fraGallery.ScrollTop = 0
    For i = 1 To count
        If IsPicked(i) Then ActiveIndex = i
    Next i
    PPTK_UpdateInfo
End Sub

Private Function PPTK_Plural(ByVal n As Long, ByVal one As String, ByVal many As String) As String
    If n = 1 Then PPTK_Plural = one Else PPTK_Plural = many
End Function

Private Sub PPTK_UpdateInfo()
    Dim rgbValue As Long
    Dim total As Long
    Dim fills As Long
    Dim lines As Long
    Dim fonts As Long
    Dim slideList As String
    Dim parts As String
    Dim picks As Long
    Dim uses As Long
    Dim slides As String

    If ActiveIndex >= 1 And ActiveIndex <= FoundCount Then
        PPTK_ColorInfo ActiveIndex, rgbValue, total, fills, lines, fonts, slideList
        parts = ""
        If fills > 0 Then parts = parts & ", fill " & fills
        If lines > 0 Then parts = parts & ", outline " & lines
        If fonts > 0 Then parts = parts & ", font " & fonts
        lblHex.caption = PPTK_ColorHex(rgbValue)
        lblUses.caption = total & PPTK_Plural(total, " use", " uses") & "  (" & Mid$(parts, 3) & ")"
        lblSlides.caption = PPTK_Plural(PPTK_ListCount(slideList), "Slide ", "Slides ") & slideList
        lblSlides.ForeColor = Ink
    Else
        lblHex.caption = ""
        lblUses.caption = ""
        If FoundCount > 0 Then
            lblSlides.caption = "Click colors to see where they are used and to choose them for replacing."
        Else
            lblSlides.caption = ""
        End If
        lblSlides.ForeColor = Muted
    End If

    slides = PPTK_ColorSlidesFor(PPTK_PickedCsv(picks), uses)
    If picks = 0 Then
        lblSelection.caption = ""
    Else
        lblSelection.caption = picks & PPTK_Plural(picks, " color", " colors") & " selected  " & ChrW$(&HB7) & "  " & _
            uses & PPTK_Plural(uses, " use", " uses") & " on " & PPTK_ListCount(slides) & PPTK_Plural(PPTK_ListCount(slides), " slide", " slides")
    End If
    PPTK_UpdateApply
End Sub

Private Function PPTK_ListCount(ByVal slideList As String) As Long
    If Len(slideList) = 0 Then PPTK_ListCount = 0 Else PPTK_ListCount = UBound(Split(slideList, ",")) + 1
End Function

Private Function PPTK_PickedCsv(ByRef picks As Long) As String
    Dim i As Long
    Dim csv As String
    picks = 0
    csv = ""
    For i = 1 To FoundCount
        If IsPicked(i) Then
            csv = csv & "," & FoundRgb(i)
            picks = picks + 1
        End If
    Next i
    If Len(csv) > 0 Then PPTK_PickedCsv = Mid$(csv, 2) Else PPTK_PickedCsv = ""
End Function

' ---------- new color ----------

Private Sub txtHex_Change()
    If Building Then Exit Sub
    PPTK_UpdateTarget
End Sub

Private Sub PPTK_UpdateTarget()
    Dim rgbValue As Long
    Dim text As String
    text = Trim$(txtHex.text)
    If PPTK_ParseHex(text, rgbValue) Then
        lblPreview.BackColor = rgbValue
        lblHexState.ForeColor = Muted
        lblHexState.caption = "New color " & PPTK_ColorHex(rgbValue)
    Else
        lblPreview.BackColor = Surface
        If Len(text) <= 1 Then
            lblHexState.ForeColor = Dim1
            lblHexState.caption = "Type #RRGGBB or pick a color below"
        Else
            lblHexState.ForeColor = ErrorRed
            lblHexState.caption = "Not a HEX color. Use #RRGGBB."
        End If
    End If
    PPTK_UpdateApply
End Sub

Private Sub PPTK_UpdateApply()
    Dim picks As Long
    Dim rgbValue As Long
    If ApplyBox Is Nothing Then Exit Sub
    PPTK_PickedCsv picks
    ApplyReady = (picks > 0 And PPTK_ParseHex(txtHex.text, rgbValue) And ScopeOk)
    If ApplyReady Then
        ApplyBox.BackColor = Accent
        ApplyBox.BorderColor = Accent
        ApplyText.ForeColor = Bg
    Else
        ApplyBox.BackColor = Surface
        ApplyBox.BorderColor = Rule
        ApplyText.ForeColor = Dim1
    End If
End Sub

' ---------- apply ----------

Public Sub PPTK_Apply(Optional ByVal confirmed As Boolean = False)
    Dim picks As Long
    Dim csv As String
    Dim targetRgb As Long
    Dim uses As Long
    Dim slides As String
    Dim sourcesText As String
    Dim typesText As String
    Dim i As Long
    Dim answer As VbMsgBoxResult
    Dim backupPath As String
    Dim changed As Long
    Dim changedSlides As String
    Dim failed As Long
    Dim summary As String

    If Not ScopeOk Then Exit Sub
    If Not PPTK_ParseHex(txtHex.text, targetRgb) Then
        MsgBox "Enter the new color as #RRGGBB, for example #795DE5.", vbExclamation, "PPTK"
        Exit Sub
    End If
    csv = ""
    sourcesText = ""
    picks = 0
    For i = 1 To FoundCount
        If IsPicked(i) And FoundRgb(i) <> targetRgb Then
            csv = csv & "," & FoundRgb(i)
            sourcesText = sourcesText & ", " & PPTK_ColorHex(FoundRgb(i))
            picks = picks + 1
        End If
    Next i
    If picks = 0 Then
        MsgBox "Select at least one color that differs from the new color.", vbInformation, "PPTK"
        Exit Sub
    End If
    csv = Mid$(csv, 2)
    sourcesText = Mid$(sourcesText, 3)
    slides = PPTK_ColorSlidesFor(csv, uses)
    typesText = ""
    If TypeOn(1) Then typesText = typesText & ", fill"
    If TypeOn(2) Then typesText = typesText & ", outline"
    If TypeOn(3) Then typesText = typesText & ", font"
    typesText = Mid$(typesText, 3)

    If Not confirmed Then
        answer = MsgBox("Replace " & picks & PPTK_Plural(picks, " color", " colors") & " with " & PPTK_ColorHex(targetRgb) & "?" & vbCr & vbCr & _
            "Colors: " & sourcesText & vbCr & _
            "Color types: " & typesText & vbCr & _
            uses & PPTK_Plural(uses, " use", " uses") & " on " & PPTK_ListCount(slides) & PPTK_Plural(PPTK_ListCount(slides), " slide", " slides") & " (" & slides & ")" & vbCr & vbCr & _
            "A backup copy of the presentation is saved first.", _
            vbOKCancel + vbQuestion + vbDefaultButton2, "PPTK - Replace Colors")
        If answer <> vbOK Then Exit Sub
    End If

    backupPath = PPTK_SaveBackupCopy(ActivePresentation, "PPTK Color Backups", "before color replace")
    If Len(backupPath) = 0 And Not confirmed Then
        answer = MsgBox("A backup copy could not be saved." & vbCr & "Replace anyway? You can still undo with Ctrl+Z.", vbYesNo + vbExclamation + vbDefaultButton2, "PPTK")
        If answer <> vbYes Then Exit Sub
    End If

    Me.MousePointer = fmMousePointerHourGlass
    changed = PPTK_ColorReplace(ScopeFrom, ScopeTo, TypeOn(1), TypeOn(2), TypeOn(3), csv, targetRgb, changedSlides, failed)
    Me.MousePointer = fmMousePointerDefault

    For i = 1 To FoundCount
        IsPicked(i) = False
    Next i
    KeepPicked = ","
    PPTK_Rescan
    For i = 1 To FoundCount
        If FoundRgb(i) = targetRgb Then
            IsPicked(i) = True
            Rings(i).Visible = True
            ActiveIndex = i
        End If
    Next i
    PPTK_UpdateInfo

    summary = "Replaced " & changed & PPTK_Plural(changed, " use", " uses") & " with " & PPTK_ColorHex(targetRgb)
    If Len(changedSlides) > 0 Then summary = summary & " on " & PPTK_Plural(PPTK_ListCount(changedSlides), "slide ", "slides ") & changedSlides
    summary = summary & "."
    If failed > 0 Then summary = summary & vbCr & failed & " matching " & PPTK_Plural(failed, "property", "properties") & " could not be changed and " & PPTK_Plural(failed, "was", "were") & " left as is."
    If Len(backupPath) > 0 Then
        summary = summary & vbCr & vbCr & "Backup copy: " & backupPath
    Else
        summary = summary & vbCr & vbCr & "No backup copy was saved."
    End If
    summary = summary & vbCr & "Press Ctrl+Z after closing this dialog to undo."
    If Not confirmed Then MsgBox summary, vbInformation, "PPTK - Replace Colors"
    LastSummary = summary
End Sub

' ---------- test hooks (read-only state for automated tests) ----------

Public Property Get LastSummaryText() As String
    LastSummaryText = LastSummary
End Property

Public Function PPTK_State() As String
    Dim picks As Long
    PPTK_State = "scopeOk=" & ScopeOk & " from=" & ScopeFrom & " to=" & ScopeTo & " found=" & FoundCount & _
        " picked=" & PPTK_PickedCsv(picks) & " apply=" & ApplyReady & _
        " info=[" & lblHex.caption & " " & lblUses.caption & " / " & lblSlides.caption & "]" & _
        " sel=[" & lblSelection.caption & "] hex=[" & lblHexState.caption & "] empty=[" & IIf(lblEmpty.Visible, lblEmpty.caption, "") & "] skip=[" & lblSkip.caption & "]"
End Function

Public Sub PPTK_TestSet(ByVal what As String, ByVal value As String)
    Select Case what
        Case "scope.all": PPTK_SwatchClick "scope", 1
        Case "scope.current": PPTK_SwatchClick "scope", 2
        Case "scope.range": PPTK_SwatchClick "scope", 3
        Case "from": txtFrom.text = value
        Case "to": txtTo.text = value
        Case "fill": If TypeOn(1) <> (value = "1") Then PPTK_SwatchClick "type", 1
        Case "line": If TypeOn(2) <> (value = "1") Then PPTK_SwatchClick "type", 2
        Case "font": If TypeOn(3) <> (value = "1") Then PPTK_SwatchClick "type", 3
        Case "hex": txtHex.text = value
    End Select
End Sub

Public Function PPTK_TestIndexOf(ByVal hexText As String) As Long
    Dim i As Long
    For i = 1 To FoundCount
        If PPTK_ColorHex(FoundRgb(i)) = UCase$(hexText) Then PPTK_TestIndexOf = i
    Next i
End Function


